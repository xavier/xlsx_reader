defmodule XlsxReader.Parsers.StylesParser do
  @moduledoc false

  # Parses SpreadsheetML style definitions.
  #
  # It extracts the relevant subset of style definitions in order to build
  # a `style_types` array which is used to look up the cell value format
  # for type conversions and the cell background color.

  @behaviour Saxy.Handler

  alias XlsxReader.{Array, Styles}
  alias XlsxReader.Parsers.Utils

  # Default indexed color palette (ECMA-376 18.8.27 `indexedColors`).
  #
  # Indexes 64 (system foreground) and 65 (system background) are deliberately
  # omitted: they depend on the system theme and have no RGB value of their own.
  @default_indexed_colors ~w(
                            00000000 00FFFFFF 00FF0000 0000FF00 000000FF 00FFFF00 00FF00FF 0000FFFF
                            00000000 00FFFFFF 00FF0000 0000FF00 000000FF 00FFFF00 00FF00FF 0000FFFF
                            00800000 00008000 00000080 00808000 00800080 00008080 00C0C0C0 00808080
                            009999FF 00993366 00FFFFCC 00CCFFFF 00660066 00FF8080 000066CC 00CCCCFF
                            00000080 00FF00FF 00FFFF00 0000FFFF 00800080 00800000 00008080 000000FF
                            0000CCFF 00CCFFFF 00CCFFCC 00FFFF99 0099CCFF 00FF99CC 00CC99FF 00FFCC99
                            003366FF 0033CCCC 0099CC00 00FFCC00 00FF9900 00FF6600 00666699 00969696
                            00003366 00339966 00003300 00333300 00993300 00993366 00333399 00333333
                            )
                          |> Enum.with_index()
                          |> Map.new(fn {rgb, index} -> {index, rgb} end)

  defmodule State do
    @moduledoc false

    # `fills` holds the background color of each fill indexed by `fillId`.
    # It is accumulated as a reversed list within the `fills` element and
    # turned into a map once all the fills have been collected.

    defstruct current_fill: nil,
              fills: %{},
              indexed_colors: [],
              style_types: [],
              custom_formats: %{},
              expect_fills: false,
              expect_indexed_colors: false,
              expect_xfs: false,
              supported_custom_formats: []
  end

  def parse(xml, supported_custom_formats \\ []) do
    with {:ok, state} <-
           Saxy.parse_string(xml, __MODULE__, %State{
             supported_custom_formats: supported_custom_formats
           }) do
      {:ok, state.style_types, state.custom_formats}
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, %State{} = state) do
    # The indexed color palette may be redefined at the very end of the
    # stylesheet: the colors referenced by the cell formats can only be
    # resolved once the whole document has been parsed.
    {:ok, %{state | style_types: build_style_types(state)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"numFmt", attributes}, %State{} = state) do
    num_fmt_id = Utils.get_attribute(attributes, "numFmtId")
    format_code = Utils.get_attribute(attributes, "formatCode")
    {:ok, %{state | custom_formats: Map.put(state.custom_formats, num_fmt_id, format_code)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"fills", _attributes}, %State{} = state) do
    {:ok, %{state | expect_fills: true, fills: []}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"fill", _attributes}, %State{expect_fills: true} = state) do
    {:ok, %{state | current_fill: %{pattern_type: nil, fg_color: nil, bg_color: nil}}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"patternFill", attributes},
        %State{current_fill: %{} = current_fill} = state
      ) do
    pattern_type = Utils.get_attribute(attributes, "patternType")
    {:ok, %{state | current_fill: %{current_fill | pattern_type: pattern_type}}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"fgColor", attributes},
        %State{current_fill: %{} = current_fill} = state
      ) do
    {:ok, %{state | current_fill: %{current_fill | fg_color: color_reference(attributes)}}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"bgColor", attributes},
        %State{current_fill: %{} = current_fill} = state
      ) do
    {:ok, %{state | current_fill: %{current_fill | bg_color: color_reference(attributes)}}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"indexedColors", _attributes}, %State{} = state) do
    {:ok, %{state | expect_indexed_colors: true}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"rgbColor", attributes},
        %State{expect_indexed_colors: true} = state
      ) do
    rgb = Utils.get_attribute(attributes, "rgb")
    {:ok, %{state | indexed_colors: [rgb | state.indexed_colors]}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"cellXfs", _attributes}, %State{} = state) do
    {:ok, %{state | expect_xfs: true}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"xf", attributes}, %State{expect_xfs: true} = state) do
    num_fmt_id = Utils.get_attribute(attributes, "numFmtId")
    fill_id = Utils.get_attribute(attributes, "fillId")

    style_type = %{
      num_fmt:
        Styles.get_style_type(
          num_fmt_id,
          state.custom_formats,
          state.supported_custom_formats
        ),
      bg_color: lookup_fill(state.fills, fill_id)
    }

    {:ok, %{state | style_types: [style_type | state.style_types]}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "fill", %State{current_fill: %{} = current_fill} = state) do
    background_color = pattern_background_color(current_fill)
    {:ok, %{state | current_fill: nil, fills: [background_color | state.fills]}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "fills", %State{expect_fills: true} = state) do
    {:ok, %{state | expect_fills: false, fills: index_fills(state.fills)}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "indexedColors", %State{} = state) do
    {:ok,
     %{
       state
       | expect_indexed_colors: false,
         indexed_colors: Enum.reverse(state.indexed_colors)
     }}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "cellXfs", %State{} = state) do
    {:ok, %{state | expect_xfs: false}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, %State{} = state) do
    {:ok, state}
  end

  ##

  defp build_style_types(%State{} = state) do
    indexed_colors = indexed_colors(state.indexed_colors)

    state.style_types
    |> Enum.reverse()
    |> Enum.map(fn style_type ->
      %{style_type | bg_color: resolve_color(style_type.bg_color, indexed_colors)}
    end)
    |> Array.from_list()
  end

  # A stylesheet may redefine part or all of the indexed color palette
  defp indexed_colors([]), do: @default_indexed_colors

  defp indexed_colors(colors),
    do: Map.merge(@default_indexed_colors, index_by_position(colors))

  # Without a pattern type nothing is painted, whatever colors the fill
  # declares. In a `solid` pattern fill the foreground color is the color of
  # the cell background and `bgColor` is unused.
  #
  # Any other pattern is drawn in `fgColor` over a background painted in
  # `bgColor`, which is reported as-is. Such a cell has no single background
  # color and applications may well display another one: LibreOffice, for
  # instance, blends both colors. Gradient fills are not supported.
  defp pattern_background_color(%{pattern_type: nil}), do: nil
  defp pattern_background_color(%{pattern_type: "none"}), do: nil
  defp pattern_background_color(%{pattern_type: "solid", fg_color: fg_color}), do: fg_color
  defp pattern_background_color(%{bg_color: bg_color}), do: bg_color

  defp index_fills(fills),
    do: fills |> Enum.reverse() |> index_by_position()

  defp index_by_position(list) do
    list
    |> Enum.with_index()
    |> Map.new(fn {value, index} -> {index, value} end)
  end

  defp lookup_fill(_fills, nil), do: nil
  defp lookup_fill(fills, fill_id), do: Map.get(fills, String.to_integer(fill_id))

  # `auto` and `theme` colors cannot be resolved from the stylesheet alone
  defp color_reference(attributes) do
    case {Utils.get_attribute(attributes, "rgb"), Utils.get_attribute(attributes, "indexed")} do
      {nil, nil} -> nil
      {nil, indexed} -> {:indexed, String.to_integer(indexed)}
      {rgb, _indexed} -> {:rgb, rgb}
    end
  end

  defp resolve_color(nil, _indexed_colors), do: nil
  defp resolve_color({:rgb, rgb}, _indexed_colors), do: rgb
  defp resolve_color({:indexed, index}, indexed_colors), do: Map.get(indexed_colors, index)
end

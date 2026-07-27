defmodule XlsxReader.Parsers.StylesParser do
  @moduledoc false

  # Parses SpreadsheetML style definitions.
  #
  # It extracts the relevant subset of style definitions in order to build
  # a `style_types` array which is used to look up the cell value format
  # for type conversions, as well as the fill (background color) of each style.

  @behaviour Saxy.Handler

  alias XlsxReader.{Array, Styles}
  alias XlsxReader.Parsers.Utils

  defmodule State do
    @moduledoc false
    defstruct pointer: nil,
              style_types: [],
              custom_formats: %{},
              fills: %{},
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
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"numFmt", attributes}, %State{} = state) do
    num_fmt_id = Utils.get_attribute(attributes, "numFmtId")
    format_code = Utils.get_attribute(attributes, "formatCode")
    {:ok, %{state | custom_formats: Map.put(state.custom_formats, num_fmt_id, format_code)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"fills", _attributes}, %State{} = state) do
    {:ok, %{state | pointer: {:collect_fills, {0, %{}}}}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"fill", _attributes},
        %State{pointer: {:collect_fills, acc}} = state
      ) do
    {:ok, %{state | pointer: {:collect_fill, nil, acc}}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"cellXfs", _attributes}, %State{} = state) do
    {:ok, %{state | pointer: :collect_xfs}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"xf", attributes}, %State{pointer: :collect_xfs} = state) do
    num_fmt_id = Utils.get_attribute(attributes, "numFmtId")
    fill_id = Utils.get_attribute(attributes, "fillId")

    style_type = %{
      num_fmt:
        Styles.get_style_type(
          num_fmt_id,
          state.custom_formats,
          state.supported_custom_formats
        ),
      fill: fill_id && state.fills[String.to_integer(fill_id)]
    }

    {:ok, %{state | style_types: [style_type | state.style_types]}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"bgColor", [{"rgb", color}]},
        %State{pointer: {:collect_fill, fill, fills}} = state
      ) do
    {:ok, %{state | pointer: {:collect_fill, Map.put(fill || %{}, :bg_color, color), fills}}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"bgColor", [{"indexed", "64"}]},
        %State{pointer: {:collect_fill, fill, fills}} = state
      ) do
    {:ok,
     %{state | pointer: {:collect_fill, Map.put(fill || %{}, :bg_color, :use_fg_color), fills}}}
  end

  @impl Saxy.Handler
  def handle_event(
        :start_element,
        {"fgColor", [{"rgb", color}]},
        %State{pointer: {:collect_fill, fill, fills}} = state
      ) do
    {:ok, %{state | pointer: {:collect_fill, Map.put(fill || %{}, :fg_color, color), fills}}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "cellXfs", %State{} = state) do
    {:ok, %{state | pointer: nil, style_types: Array.from_list(Enum.reverse(state.style_types))}}
  end

  @impl Saxy.Handler
  def handle_event(
        :end_element,
        "fill",
        %State{pointer: {:collect_fill, fill, {i, fills}}} = state
      ) do
    # `bgColor indexed="64"` means the actual background color is the pattern's
    # foreground color (the case for solid pattern fills).
    fill =
      case fill do
        %{bg_color: :use_fg_color, fg_color: color} -> Map.put(fill, :bg_color, color)
        %{bg_color: :use_fg_color} -> Map.drop(fill, [:bg_color])
        _ -> fill
      end

    {:ok, %{state | pointer: {:collect_fills, {i + 1, Map.put(fills, i, fill)}}}}
  end

  @impl Saxy.Handler
  def handle_event(
        :end_element,
        "fills",
        %State{pointer: {:collect_fills, {_count, fills}}} = state
      ) do
    {:ok, %{state | pointer: nil, fills: fills}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, %State{} = state) do
    {:ok, state}
  end
end

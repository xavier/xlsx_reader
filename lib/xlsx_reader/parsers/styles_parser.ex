defmodule XlsxReader.Parsers.StylesParser do
  @moduledoc false

  # Parses SpreadsheetML style definitions.
  #
  # It extracts the relevant subset of style definitions in order to build
  # a `style_types` array which is used to look up the cell value format
  # for type conversions.

  @behaviour Saxy.Handler

  alias XlsxReader.{Array, Styles}
  alias XlsxReader.Parsers.Utils

  defmodule State do
    @moduledoc false
    defstruct collect_xf: false,
              style_types: [],
              custom_formats: %{},
              supported_custom_formats: []
  end

  def parse(xml, supported_custom_formats \\ []) do
    xml = Utils.strip_leading_bom(xml)

    with {:ok, state} <-
           Saxy.parse_string(xml, __MODULE__, %State{
             supported_custom_formats: supported_custom_formats
           }) do
      {:ok, state.style_types, state.custom_formats}
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"numFmt", attributes}, state) do
    num_fmt_id = Utils.get_attribute(attributes, "numFmtId")
    format_code = Utils.get_attribute(attributes, "formatCode")
    {:ok, %{state | custom_formats: Map.put(state.custom_formats, num_fmt_id, format_code)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"cellXfs", _attributes}, state) do
    {:ok, %{state | collect_xf: true}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"xf", attributes}, %{collect_xf: true} = state) do
    num_fmt_id = Utils.get_attribute(attributes, "numFmtId")

    {:ok,
     %{
       state
       | style_types: [
           Styles.get_style_type(
             num_fmt_id,
             state.custom_formats,
             state.supported_custom_formats
           )
           | state.style_types
         ]
     }}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "cellXfs", state) do
    {:ok,
     %{state | collect_xf: false, style_types: Array.from_list(Enum.reverse(state.style_types))}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, state) do
    {:ok, state}
  end
end

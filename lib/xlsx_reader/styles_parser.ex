defmodule XlsxReader.StylesParser do
  @behaviour Saxy.Handler

  alias XlsxReader.Styles

  def parse(xml) do
    Saxy.parse_string(xml, __MODULE__, %{collect_xf: false, style_types: [], custom_formats: %{}})
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, Map.take(state, [:style_types, :custom_formats])}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"numFmt", attributes}, state) do
    num_fmt_id = get_attribute(attributes, "numFmtId")
    format_code = get_attribute(attributes, "formatCode")
    {:ok, %{state | custom_formats: Map.put(state.custom_formats, num_fmt_id, format_code)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"cellXfs", attributes}, state) do
    {:ok, %{state | collect_xf: true}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"xf", attributes}, %{collect_xf: true} = state) do
    num_fmt_id = get_attribute(attributes, "numFmtId")

    {:ok,
     %{
       state
       | style_types: [
           Styles.get_style_type(num_fmt_id, state.custom_formats) | state.style_types
         ]
     }}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "cellXfs", state) do
    {:ok, %{state | collect_xf: false, style_types: Enum.reverse(state.style_types)}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, state) do
    {:ok, state}
  end

  ##

  defp get_attribute(attributes, name, default \\ nil)
  defp get_attribute([], _name, default), do: default
  defp get_attribute([{name, value} | _], name, _default), do: value
  defp get_attribute([_ | rest], name, default), do: get_attribute(rest, name, default)
end

defmodule XlsxReader.WorksheetParser do
  @behaviour Saxy.Handler

  def parse(xml, shared_strings) do
    Saxy.parse_string(xml, __MODULE__, %{
      shared_strings: shared_strings,
      rows: [],
      row: nil,
      cell_ref: nil,
      cell_type: nil,
      value_type: nil,
      value: nil
    })
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, state.rows}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"row", attributes}, state) do
    {:ok, %{state | row: []}}
  end

  def handle_event(:start_element, {"c", attributes}, state) do
    {:ok, Map.merge(state, map_cell_attributes(attributes))}
  end

  def handle_event(:start_element, {"v", _attributes}, state) do
    {:ok, %{state | value_type: :value, value: :expect_chars}}
  end

  def handle_event(:start_element, {"f", _attributes}, state) do
    {:ok, %{state | value_type: :formula, value: :expect_chars}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "c", state) do
    {:ok,
     %{
       state
       | row: [format_current_cell_value(state) | state.row],
         cell_ref: nil,
         cell_type: nil,
         value_type: nil,
         value: nil
     }}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "row", state) do
    {:ok, %{state | rows: [Enum.reverse(state.row) | state.rows], row: nil}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sheetData", state) do
    {:ok, %{state | rows: Enum.reverse(state.rows)}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, chars, %{value: :expect_chars} = state) do
    case state.value_type do
      :value ->
        {:ok, %{state | value: chars}}

        :formula
        # TODO
        {:ok, %{state | value: chars}}

      _ ->
        {:ok, state}
    end
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, state) do
    {:ok, state}
  end

  ##

  @attributes_mapping %{
    "r" => :cell_ref,
    # "s" => :cell_style,
    "t" => :cell_type
  }

  defp map_cell_attributes(attributes) do
    Enum.reduce(attributes, %{}, fn {name, value}, acc ->
      case Map.fetch(@attributes_mapping, name) do
        {:ok, key} ->
          Map.put(acc, key, value)

        :error ->
          acc
      end
    end)
  end

  defp format_current_cell_value(state) do
    case {state.cell_type, state.value} do
      {"s", value} ->
        Enum.at(state.shared_strings, String.to_integer(value))

      {nil, value} ->
        to_string(value)

      {_, value} ->
        to_string(value)
    end
  end
end

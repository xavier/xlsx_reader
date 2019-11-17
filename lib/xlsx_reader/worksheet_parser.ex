defmodule XlsxReader.WorksheetParser do
  @behaviour Saxy.Handler

  alias XlsxReader.Conversion

  defmodule State do
    defstruct workbook: nil,
              rows: [],
              row: nil,
              cell_ref: nil,
              cell_type: nil,
              value_type: nil,
              value: nil
  end

  def parse(xml, workbook) do
    Saxy.parse_string(xml, __MODULE__, %State{workbook: workbook})
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
  def handle_event(:start_element, {"row", _attributes}, state) do
    {:ok, %{state | row: []}}
  end

  def handle_event(:start_element, {"c", attributes}, state) do
    {:ok, Map.merge(state, map_cell_attributes(attributes))}
  end

  def handle_event(:start_element, {"v", _attributes}, state) do
    {:ok, %{state | value_type: :value, value: :expect_chars}}
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
    "s" => :cell_style,
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
    style_type = Enum.at(state.workbook.style_types, String.to_integer(state.cell_style))

    case {state.cell_type, style_type, state.value} do
      {_, _, nil} ->
        ""

      {"s", _, value} ->
        Enum.at(state.workbook.shared_strings, String.to_integer(value))

      {_, :percentage, value} ->
        {:ok, number} = Conversion.to_number(value)
        number * 100

      {nil, :date, value} ->
        {:ok, date} = Conversion.to_date(value, state.workbook.base_date)
        date

      {nil, :time, value} ->
        {:ok, date_time} = Conversion.to_date_time(value, state.workbook.base_date)
        date_time

      {nil, :date_time, value} ->
        {:ok, date_time} = Conversion.to_date_time(value, state.workbook.base_date)
        date_time

      {"b", _, "1"} ->
        true

      {"b", _, "0"} ->
        false

      {nil, _, value} ->
        to_string(value)

      {_, _, value} ->
        to_string(value)
    end
  end
end

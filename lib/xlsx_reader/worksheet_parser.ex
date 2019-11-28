defmodule XlsxReader.WorksheetParser do
  @moduledoc false

  # Parses SpreadsheetML worksheets.

  @behaviour Saxy.Handler

  alias XlsxReader.{Conversion, Number, ParserUtils}

  defmodule State do
    @moduledoc false
    @enforce_keys [:workbook, :type_conversion, :blank_value]
    defstruct workbook: nil,
              rows: [],
              row: nil,
              cell_ref: nil,
              cell_type: nil,
              value: nil,
              type_conversion: nil,
              blank_value: nil,
              empty_rows: nil,
              number_type: nil
  end

  @doc """
  Parse the given worksheet XML in the context of the given workbook.

  ## Options

    * `type_conversion`: boolean (default: `true`)
    * `blank_value`: placeholder value for empty cells (default: `""`)
    * `empty_rows`: include empty rows (default: `true`)
    * `number_type` - type used for numeric conversion :`Integer`, 'Decimal' or `Float` (default: `Float`)

  """
  def parse(xml, workbook, options \\ []) do
    Saxy.parse_string(xml, __MODULE__, %State{
      workbook: workbook,
      type_conversion: Keyword.get(options, :type_conversion, true),
      blank_value: Keyword.get(options, :blank_value, ""),
      empty_rows: Keyword.get(options, :empty_rows, true),
      number_type: Keyword.get(options, :number_type, Float)
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
  def handle_event(:start_element, {"row", _attributes}, state) do
    {:ok, %{state | row: []}}
  end

  def handle_event(:start_element, {"c", attributes}, state) do
    {:ok, Map.merge(state, extract_cell_attributes(attributes))}
  end

  def handle_event(:start_element, {"v", _attributes}, state) do
    {:ok, expect_value(state)}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "c", state) do
    {:ok, add_cell(state)}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "row", state) do
    if skip_row?(state) do
      {:ok, skip_row(state)}
    else
      {:ok, emit_row(state)}
    end
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sheetData", state) do
    {:ok, restore_rows_order(state)}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, chars, %{value: :expect_chars} = state) do
    {:ok, store_value(state, chars)}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, state) do
    {:ok, state}
  end

  ##

  ## State machine

  defp expect_value(state) do
    %{state | value: :expect_chars}
  end

  defp store_value(state, value) do
    %{state | value: value}
  end

  defp add_cell(state) do
    %{
      state
      | row: [convert_current_cell_value(state) | state.row],
        cell_ref: nil,
        cell_type: nil,
        value: nil
    }
  end

  defp skip_row?(state) do
    !state.empty_rows && empty_row?(state)
  end

  defp empty_row?(state) do
    Enum.all?(state.row, fn value -> value == state.blank_value end)
  end

  defp skip_row(state) do
    %{state | row: nil}
  end

  defp emit_row(state) do
    %{state | row: nil, rows: [Enum.reverse(state.row) | state.rows]}
  end

  defp restore_rows_order(state) do
    %{state | rows: Enum.reverse(state.rows)}
  end

  ## Cell format handling

  @cell_attributes_mapping %{
    "r" => :cell_ref,
    "s" => :cell_style,
    "t" => :cell_type
  }

  defp extract_cell_attributes(attributes) do
    ParserUtils.map_attributes(attributes, @cell_attributes_mapping)
  end

  defp convert_current_cell_value(%State{type_conversion: false} = state) do
    case {state.cell_type, state.value} do
      {_, nil} ->
        state.blank_value

      {"s", value} ->
        lookup_shared_string(state, value)

      {_, value} ->
        value
    end
  end

  defp convert_current_cell_value(%State{type_conversion: true} = state) do
    style_type = Enum.at(state.workbook.style_types, String.to_integer(state.cell_style))

    case {state.cell_type, style_type, state.value} do
      {_, _, nil} ->
        state.blank_value

      {"s", _, value} ->
        lookup_shared_string(state, value)

      {"b", _, value} ->
        {:ok, boolean} = Conversion.to_boolean(value)
        boolean

      {_, :percentage, value} ->
        {:ok, number} = Conversion.to_number(value, state.number_type)
        Number.multiply(number, 100)

      {nil, :date, value} ->
        {:ok, date} = Conversion.to_date(value, state.workbook.base_date)
        date

      {nil, :time, value} ->
        {:ok, date_time} = Conversion.to_date_time(value, state.workbook.base_date)
        date_time

      {nil, :date_time, value} ->
        {:ok, date_time} = Conversion.to_date_time(value, state.workbook.base_date)
        date_time

      {nil, _, value} ->
        {:ok, number} = Conversion.to_number(value, state.number_type)
        number

      {_, _, value} ->
        value
    end
  end

  defp lookup_shared_string(state, value) do
    Enum.at(state.workbook.shared_strings, String.to_integer(value))
  end
end

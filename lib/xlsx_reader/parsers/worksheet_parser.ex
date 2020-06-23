defmodule XlsxReader.Parsers.WorksheetParser do
  @moduledoc false

  # Parses SpreadsheetML worksheets.

  @behaviour Saxy.Handler

  alias XlsxReader.{CellReference, Conversion, Number}
  alias XlsxReader.Parsers.Utils

  defmodule State do
    @moduledoc false
    @enforce_keys [:workbook, :type_conversion, :blank_value]
    defstruct workbook: nil,
              rows: [],
              row: nil,
              expected_column: nil,
              cell_ref: nil,
              cell_type: nil,
              cell_style: nil,
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
    * `number_type` - type used for numeric conversion : `String` (no conversion), `Integer`, 'Decimal' or `Float`  (default: `Float`)

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
    # Some XLSX writers (Excel, Elixlsx, …) completely omit `<c>` elements for empty cells.
    # As we build the row, we'll keep track of the expected column number and fill the blanks
    # based on the column indicated in the cell reference
    {:ok, %{state | row: [], expected_column: 1}}
  end

  def handle_event(:start_element, {"c", attributes}, state) do
    {:ok, new_cell(state, extract_cell_attributes(attributes))}
  end

  def handle_event(:start_element, {"v", _attributes}, state) do
    {:ok, expect_value(state)}
  end

  def handle_event(:start_element, {"t", _attributes}, state) do
    {:ok, expect_value(state)}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "c", state) do
    {:ok, add_cell_to_row(state)}
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

  defp new_cell(state, cell_attributes) do
    state
    |> Map.merge(cell_attributes)
    |> handle_omitted_cells()
  end

  defp expect_value(state) do
    %{state | value: :expect_chars}
  end

  defp store_value(state, value) do
    %{state | value: value}
  end

  defp add_cell_to_row(state) do
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

  ## Omitted cells

  defp handle_omitted_cells(state) do
    # Using the current cell reference and the expected column:
    # 1. fill any missing cell
    # 2. determine the next expected column
    with %{cell_ref: cell_ref} when not is_nil(cell_ref) <- state,
         {column, _row} <- CellReference.parse(cell_ref) do
      omitted_cells = column - state.expected_column

      state
      |> add_omitted_cells_to_row(omitted_cells)
      |> Map.put(:expected_column, column + 1)
    else
      _ ->
        state
    end
  end

  defp add_omitted_cells_to_row(state, n) do
    %{state | row: prepend_n_times(state.row, state.blank_value, n)}
  end

  defp prepend_n_times(list, _value, 0), do: list
  defp prepend_n_times(list, value, n), do: prepend_n_times([value | list], value, n - 1)

  ## Cell format handling

  @cell_attributes_mapping %{
    "r" => :cell_ref,
    "s" => :cell_style,
    "t" => :cell_type
  }

  defp extract_cell_attributes(attributes) do
    Utils.map_attributes(attributes, @cell_attributes_mapping)
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

  # credo:disable-for-lines:54 Credo.Check.Refactor.CyclomaticComplexity
  defp convert_current_cell_value(%State{type_conversion: true} = state) do
    style_type = lookup_current_cell_style_type(state)

    case {state.cell_type, style_type, state.value} do
      # Blank

      {_, _, value} when is_nil(value) or value == "" ->
        state.blank_value

      # Strings

      {"s", _, value} ->
        lookup_shared_string(state, value)

      {"inlineStr", _, value} ->
        value

      {nil, :string, value} ->
        value

      {"b", _, value} ->
        {:ok, boolean} = Conversion.to_boolean(value)
        boolean

      # Numbers

      {"n", _, value} ->
        {:ok, number} = Conversion.to_number(value, state.number_type)
        number

      {nil, :number, value} ->
        {:ok, number} = Conversion.to_number(value, state.number_type)
        number

      {_, :percentage, value} ->
        {:ok, number} = Conversion.to_number(value, state.number_type)
        Number.multiply(number, 100)

      # Dates/times

      {nil, :date, value} ->
        {:ok, date} = Conversion.to_date(value, state.workbook.base_date)
        date

      {nil, type, value} when type in [:time, :date_time] ->
        {:ok, date_time} = Conversion.to_date_time(value, state.workbook.base_date)
        date_time

      # Fall back

      {_, _, value} ->
        value
    end
  end

  defp lookup_current_cell_style_type(state) do
    if state.cell_style,
      do: lookup_index(state.workbook.style_types, state.cell_style),
      else: nil
  end

  defp lookup_shared_string(state, value) do
    lookup_index(state.workbook.shared_strings, value)
  end

  defp lookup_index(table, string_index) do
    {:ok, index} = Conversion.to_integer(string_index)
    XlsxReader.Array.lookup(table, index)
  end
end

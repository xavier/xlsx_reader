defmodule XlsxReader.Parsers.WorksheetParser do
  @moduledoc false

  # Parses SpreadsheetML worksheets.

  @behaviour Saxy.Handler

  alias XlsxReader.{Cell, CellReference, Conversion, Number}
  alias XlsxReader.Parsers.Utils

  defmodule State do
    @moduledoc false
    @enforce_keys [:workbook, :type_conversion, :blank_value]
    defstruct workbook: nil,
              rows: [],
              row: nil,
              current_row: nil,
              expected_row: 1,
              expected_column: nil,
              cell_ref: nil,
              cell_type: nil,
              cell_style: nil,
              value: nil,
              formula: nil,
              type_conversion: nil,
              blank_value: nil,
              empty_rows: nil,
              number_type: nil,
              skip_row?: nil,
              cell_data_format: :value
  end

  @doc """
  Parse the given worksheet XML in the context of the given workbook.

  ## Options

    * `type_conversion`: boolean (default: `true`)
    * `blank_value`: placeholder value for empty cells (default: `""`)
    * `empty_rows`: include empty rows (default: `true`)
    * `number_type` - type used for numeric conversion : `String` (no conversion), `Integer`, 'Decimal' or `Float`  (default: `Float`)
    * `skip_row?`: function callback that determines if a row should be skipped or not.
       Overwrites `blank_value` and `empty_rows` on the matter of skipping rows.
       Defaults to `nil` (keeping the behaviour of `blank_value` and `empty_rows`).
    * `cell_data_format`: Controls the format of the cell data. Can be `:value` (default, returns the cell value only) or `:cell` (returns instances of `XlsxReader.Cell`).

  """
  def parse(xml, workbook, options \\ []) do
    Saxy.parse_string(xml, __MODULE__, %State{
      workbook: workbook,
      type_conversion: Keyword.get(options, :type_conversion, true),
      blank_value: Keyword.get(options, :blank_value, ""),
      empty_rows: Keyword.get(options, :empty_rows, true),
      number_type: Keyword.get(options, :number_type, Float),
      skip_row?: Keyword.get(options, :skip_row?),
      cell_data_format: Keyword.get(options, :cell_data_format, :value)
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
    # Some XLSX writers (Excel, Elixlsx, …) completely omit `<row>` or `<c>` elements when empty.
    # As we build the sheet, we'll keep track of the expected row and column number and
    # fill the blanks as needed usingthe coordinates indicated in the row or cell reference.

    current_row =
      case Utils.get_attribute(attributes, "r") do
        nil ->
          state.expected_row

        value ->
          String.to_integer(value)
      end

    {:ok,
     %{
       state
       | row: [],
         current_row: current_row,
         expected_column: 1
     }}
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

  def handle_event(:start_element, {"f", _attributes}, state) do
    {:ok, expect_formula(state)}
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
  def handle_event(:characters, chars, %{value: :expect_formula} = state) do
    {:ok, store_formula(state, chars)}
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

  defp expect_formula(state) do
    %{state | value: :expect_formula}
  end

  defp store_value(state, value) do
    %{state | value: value}
  end

  defp store_formula(state, formula) do
    %{state | formula: formula}
  end

  defp add_cell_to_row(state) do
    %{
      state
      | row: [format_cell_data(state) | state.row],
        cell_ref: nil,
        cell_type: nil,
        value: nil,
        formula: nil
    }
  end

  defp skip_row?(%{skip_row?: skip_row?, row: row}) when is_function(skip_row?) do
    skip_row?.(row)
  end

  defp skip_row?(%{empty_rows: false} = state), do: empty_row?(state)

  defp skip_row?(_state), do: false

  defp empty_row?(state) do
    Enum.all?(state.row, fn value -> value == state.blank_value end)
  end

  defp skip_row(state) do
    %{state | row: nil, expected_row: state.current_row + 1}
  end

  defp emit_row(state) do
    state = handle_omitted_rows(state)

    %{
      state
      | row: nil,
        rows: [Enum.reverse(state.row) | state.rows],
        expected_row: state.current_row + 1
    }
  end

  defp restore_rows_order(state) do
    %{state | rows: Enum.reverse(state.rows)}
  end

  ## Omitted rows / cells

  defp handle_omitted_rows(%{empty_rows: true} = state) do
    omitted_rows = state.current_row - state.expected_row

    if omitted_rows > 0 do
      blank_row = List.duplicate(state.blank_value, length(state.row))
      %{state | rows: prepend_n_times(state.rows, blank_row, omitted_rows)}
    else
      state
    end
  end

  defp handle_omitted_rows(state), do: state

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

  defp prepend_n_times(list, value, n) when n > 0,
    do: prepend_n_times([value | list], value, n - 1)

  ## Cell format handling

  @cell_attributes_mapping %{
    "r" => :cell_ref,
    "s" => :cell_style,
    "t" => :cell_type
  }

  defp extract_cell_attributes(attributes) do
    Utils.map_attributes(attributes, @cell_attributes_mapping)
  end

  defp format_cell_data(state) do
    value = convert_current_cell_value(state)

    cell_data = %Cell{
      value: value,
      formula: state.formula,
      ref: state.cell_ref
    }

    case state.cell_data_format do
      :cell -> cell_data
      :value -> value
      _ -> value
    end
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

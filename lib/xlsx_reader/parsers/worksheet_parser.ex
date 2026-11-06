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
              shared_formula_index: nil,
              shared_formulas: :array.new(),
              type_conversion: nil,
              blank_value: nil,
              empty_rows: nil,
              number_type: nil,
              skip_row?: nil,
              cell_data_format: :value,
              max_rows: :infinity,
              max_columns: :infinity,
              max_cells: :infinity,
              cell_count: 0
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
    * `max_rows`: maximum row number allowed in the sheet (default: `:infinity`)
    * `max_columns`: maximum column number allowed in the sheet (default: `:infinity`)
    * `max_cells`: maximum number of cells in the parsed sheet, including blank cells
       created to fill omitted rows and cells (default: `:infinity`)
  """
  def parse(xml, workbook, options \\ []) do
    xml
    |> Saxy.parse_string(__MODULE__, %State{
      workbook: workbook,
      type_conversion: Keyword.get(options, :type_conversion, true),
      blank_value: Keyword.get(options, :blank_value, ""),
      empty_rows: Keyword.get(options, :empty_rows, true),
      number_type: Keyword.get(options, :number_type, Float),
      skip_row?: Keyword.get(options, :skip_row?),
      cell_data_format: Keyword.get(options, :cell_data_format, :value),
      max_rows: Keyword.get(options, :max_rows, :infinity),
      max_columns: Keyword.get(options, :max_columns, :infinity),
      max_cells: Keyword.get(options, :max_cells, :infinity)
    })
    |> case do
      {:ok, {:limit_exceeded, reason}} -> {:error, reason}
      result -> result
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, %State{} = state) do
    {:ok, state.rows}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"row", attributes}, %State{} = state) do
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

    if exceeds?(current_row, state.max_rows) do
      limit_exceeded("row")
    else
      {:ok,
       %{
         state
         | row: [],
           current_row: current_row,
           expected_column: 1
       }}
    end
  end

  def handle_event(:start_element, {"c", attributes}, %State{} = state) do
    new_cell(state, attributes)
  end

  def handle_event(:start_element, {"v", _attributes}, %State{} = state) do
    {:ok, expect_value(state)}
  end

  def handle_event(:start_element, {"t", _attributes}, %State{} = state) do
    {:ok, expect_value(state)}
  end

  def handle_event(:end_element, "v", %State{value: :expect_chars} = state) do
    # Value element was empty, i.e. the sentinel value wasn't overridden by the :characters handler
    {:ok, store_value(state, "")}
  end

  def handle_event(:start_element, {"f", attributes}, %State{} = state) do
    type = Utils.get_attribute(attributes, "t")
    ref = Utils.get_attribute(attributes, "ref")

    case {type, ref} do
      {"shared", ref} when is_binary(ref) ->
        string_index = Utils.get_attribute(attributes, "si")
        {:ok, expect_shared_formula(state, string_index)}

      _ ->
        {:ok, expect_formula(state)}
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "c", %State{} = state) do
    {:ok, add_cell_to_row(state)}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "f", %State{formula: nil} = state) do
    formula = lookup_shared_formula(state, state.shared_formula_index)
    {:ok, store_formula(state, formula)}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "row", %State{} = state) do
    if skip_row?(state) do
      {:ok, skip_row(state)}
    else
      emit_row(state)
    end
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sheetData", %State{} = state) do
    {:ok, restore_rows_order(state)}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, %State{} = state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, chars, %State{value: :expect_chars} = state) do
    {:ok, store_value(state, chars)}
  end

  @impl Saxy.Handler
  def handle_event(:characters, chars, %State{value: :expect_formula} = state) do
    {:ok, store_formula(state, chars)}
  end

  @impl Saxy.Handler
  def handle_event(:characters, chars, %State{value: :expect_shared_formula} = state) do
    state = store_shared_formula(state, state.shared_formula_index, chars)
    {:ok, store_formula(state, chars)}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, %State{} = state) do
    {:ok, state}
  end

  ##

  ## State machine

  defp new_cell(%State{} = state, attributes) do
    state
    |> apply_cell_attributes(attributes)
    |> handle_omitted_cells()
  end

  defp apply_cell_attributes(%State{} = state, []), do: state

  defp apply_cell_attributes(%State{} = state, [{"r", value} | rest]),
    do: apply_cell_attributes(%{state | cell_ref: value}, rest)

  defp apply_cell_attributes(%State{} = state, [{"s", value} | rest]),
    do: apply_cell_attributes(%{state | cell_style: value}, rest)

  defp apply_cell_attributes(%State{} = state, [{"t", value} | rest]),
    do: apply_cell_attributes(%{state | cell_type: value}, rest)

  defp apply_cell_attributes(%State{} = state, [_ | rest]),
    do: apply_cell_attributes(state, rest)

  defp expect_value(%State{} = state) do
    %{state | value: :expect_chars}
  end

  defp expect_formula(%State{} = state) do
    %{state | value: :expect_formula}
  end

  defp expect_shared_formula(%State{} = state, string_index) do
    index = String.to_integer(string_index)
    shared_formulas = state.shared_formulas |> XlsxReader.Array.insert(index, nil)

    %{
      state
      | value: :expect_shared_formula,
        shared_formulas: shared_formulas,
        shared_formula_index: index
    }
  end

  defp store_value(%State{} = state, value) do
    %{state | value: value}
  end

  defp store_formula(%State{} = state, formula) do
    %{state | formula: formula, value: nil}
  end

  defp store_shared_formula(%State{} = state, index, formula) do
    shared_formulas = state.shared_formulas |> XlsxReader.Array.insert(index, formula)
    %{state | shared_formulas: shared_formulas}
  end

  defp add_cell_to_row(%State{} = state) do
    %{
      state
      | row: [format_cell_data(state) | state.row],
        cell_ref: nil,
        cell_style: nil,
        cell_type: nil,
        value: nil,
        formula: nil
    }
  end

  defp skip_row?(%State{skip_row?: skip_row?, row: row}) when is_function(skip_row?) do
    skip_row?.(row)
  end

  defp skip_row?(%State{empty_rows: false} = state), do: empty_row?(state)

  defp skip_row?(_state), do: false

  defp empty_row?(%State{} = state) do
    Enum.all?(state.row, fn value -> value == state.blank_value end)
  end

  defp skip_row(%State{} = state) do
    %{state | row: nil, expected_row: state.current_row + 1}
  end

  defp emit_row(%State{} = state) do
    width = length(state.row)
    cell_count = state.cell_count + width * (omitted_rows(state) + 1)

    if exceeds?(cell_count, state.max_cells) do
      limit_exceeded("cell")
    else
      state = handle_omitted_rows(state)
      row = state.row |> sanitize_row() |> Enum.reverse()

      {:ok,
       %{
         state
         | row: nil,
           rows: [row | state.rows],
           expected_row: state.current_row + 1,
           cell_count: cell_count
       }}
    end
  end

  defp sanitize_row(row) do
    Enum.map(row, fn
      # If the <c> element has no text child node, we didn't receive any :characters event
      # and the current value still contains the placeholder used by the parser
      :expect_chars -> ""
      # Otherwise assume that the row contains an actual value
      value -> value
    end)
  end

  defp restore_rows_order(%State{} = state) do
    %{state | rows: Enum.reverse(state.rows)}
  end

  ## Limits
  #
  # Limits apply to the sheet as it gets materialized (including the blank cells
  # and rows added to fill omitted ones), not to the number of XML elements:
  # a tiny sparse file can otherwise expand into billions of cells.

  defp exceeds?(_value, :infinity), do: false
  defp exceeds?(value, limit), do: value > limit

  defp limit_exceeded(kind) do
    {:stop, {:limit_exceeded, "#{kind} limit exceeded"}}
  end

  ## Omitted rows / cells

  defp omitted_rows(%State{empty_rows: true} = state),
    do: max(state.current_row - state.expected_row, 0)

  defp omitted_rows(%State{}), do: 0

  defp handle_omitted_rows(%State{} = state) do
    case omitted_rows(state) do
      0 ->
        state

      omitted_rows ->
        blank_row = List.duplicate(state.blank_value, length(state.row))
        %{state | rows: prepend_n_times(state.rows, blank_row, omitted_rows)}
    end
  end

  defp handle_omitted_cells(%State{} = state) do
    # Using the current cell reference and the expected column:
    # 1. fill any missing cell
    # 2. determine the next expected column
    with %State{cell_ref: cell_ref} when not is_nil(cell_ref) <- state,
         {column, _row} <- CellReference.parse(cell_ref) do
      cond do
        exceeds?(column, state.max_columns) ->
          limit_exceeded("column")

        exceeds?(state.cell_count + column, state.max_cells) ->
          limit_exceeded("cell")

        true ->
          omitted_cells = column - state.expected_column

          {:ok,
           state
           |> add_omitted_cells_to_row(omitted_cells)
           |> Map.put(:expected_column, column + 1)}
      end
    else
      _ ->
        {:ok, state}
    end
  end

  defp add_omitted_cells_to_row(%State{} = state, n) do
    %{state | row: prepend_n_times(state.row, state.blank_value, n)}
  end

  defp prepend_n_times(list, _value, 0), do: list

  defp prepend_n_times(list, value, n) when n > 0,
    do: prepend_n_times([value | list], value, n - 1)

  ## Cell format handling

  defp format_cell_data(%State{} = state) do
    value = convert_current_cell_value(state)

    case state.cell_data_format do
      :cell ->
        style_type = lookup_current_cell_style_type(state)

        %Cell{
          value: value,
          formula: state.formula,
          ref: state.cell_ref,
          bg_color: style_type && style_type.bg_color
        }

      :value ->
        value

      _ ->
        value
    end
  end

  defp convert_current_cell_value(%State{type_conversion: false} = state) do
    case {state.cell_type, state.value} do
      {_, value} when is_nil(value) or value == "" ->
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

    case {state.cell_type, style_type && style_type.num_fmt, state.value} do
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
        value |> Conversion.to_boolean() |> handle_conversion_error

      # Numbers

      {"n", :date, value} ->
        value |> Conversion.to_date(state.workbook.base_date) |> handle_conversion_error()

      {"n", type, value} when type in [:time, :date_time] ->
        value |> Conversion.to_date_time(state.workbook.base_date) |> handle_conversion_error()

      {"n", _, value} ->
        value |> Conversion.to_number(state.number_type) |> handle_conversion_error()

      {nil, :number, value} ->
        value |> Conversion.to_number(state.number_type) |> handle_conversion_error()

      {_, :percentage, value} ->
        case Conversion.to_number(value, state.number_type) do
          {:ok, number} ->
            Number.multiply(number, 100)

          _ ->
            "#ERROR"
        end

      # Dates/times

      {nil, :date, value} ->
        value |> Conversion.to_date(state.workbook.base_date) |> handle_conversion_error()

      {nil, type, value} when type in [:time, :date_time] ->
        value |> Conversion.to_date_time(state.workbook.base_date) |> handle_conversion_error()

      # Fall back

      {_, _, value} ->
        value
    end
  end

  defp handle_conversion_error({:ok, value}), do: value
  defp handle_conversion_error(_error), do: "#ERROR"

  defp lookup_current_cell_style_type(%State{} = state) do
    if state.cell_style,
      do: lookup_index(state.workbook.style_types, state.cell_style),
      else: nil
  end

  defp lookup_shared_string(%State{} = state, value) do
    lookup_index(state.workbook.shared_strings, value)
  end

  defp lookup_shared_formula(%State{} = state, index) do
    state.shared_formulas |> XlsxReader.Array.lookup(index)
  end

  defp lookup_index(nil, _string_index), do: nil

  defp lookup_index(table, string_index) do
    XlsxReader.Array.lookup(table, String.to_integer(string_index))
  end
end

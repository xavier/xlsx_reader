defmodule XlsxReader do
  @moduledoc """

  Opens XLSX workbooks and reads its worksheets.

  ## Example

  ```elixir
  {:ok, package} = XlsxReader.open("test.xlsx")

  XlsxReader.sheet_names(package)
  # ["Sheet 1", "Sheet 2", "Sheet 3"]

  {:ok, rows} = XlsxReader.sheet(package, "Sheet 1")
  # [
  #   ["Date", "Temperature"],
  #   [~D[2019-11-01], 8.4],
  #   [~D[2019-11-02], 7.5],
  #   ...
  # ]
  ```

  ## Sheet contents

  Sheets are loaded on-demand by `sheet/3` and `sheets/2`.

  The sheet contents is returned as a list of lists:

  ```elixir
  [
    ["A1", "B1", "C1" | _],
    ["A2", "B2", "C2" | _],
    ["A3", "B3", "C3" | _],
    | _
  ]
  ```

  The behavior of the sheet parser can be customized for each
  individual sheet, see `sheet/3`.

  ## Cell types

  This library takes a best effort approach for determining cell types.
  In order of priority, the actual type of an XLSX cell value is determined using:

  1. basic cell properties (e.g. boolean)
  2. predefined known styles (e.g. default money/date formats)
  3. introspection of the [custom format string](https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68) associated with the cell

  ### Custom formats supported by default

  * percentages
  * ISO 8601 date/time (y-m-d)
  * US date/time (m/d/y)
  * European date/time (d/m/y)

  ### Additional custom formats support

  If the spreadsheet you need to process contains some unusual cell formatting, you
  may provide hints to map format strings to a known cell type.

  The hints are given as a list of `{matcher, type}` tuples. The matcher is either a
  string or regex to match against the custom format string. The supported types are:

  * `:string`
  * `:number`
  * `:percentage`
  * `:date`
  * `:time`
  * `:date_time`
  * `:unsupported` (used for explicitly unsupported styles and formats)

  ### Conversion errors

  Cell data which could not be converted using the detected format is returned as the `"#ERROR"` placeholder.

  #### Example

  ```elixir
  [
    {"mmm yy", :date},
    {~r/mmm? yy hh:mm/, :date_time},
    {"[$CHF]0.00", :number}
  ]
  ```

  To find out what custom formats are in use in the workbook, you can inspect `package.workbook.custom_formats`:

  ```elixir
  # num_fmt_id => format string
  %{
    "0" => "General",
    "59" => "dd/mm/yyyy",
    "60" => "dd/mm/yyyy hh:mm",
    "61" => "hh:mm",
    "62" => "0.0%",
    "63" => "[$CHF]0.00"
  }
  ```

  """

  alias XlsxReader.{PackageLoader, ZipArchive}

  @typedoc """
  Source for the XLSX file: file system (`:path`) or in-memory (`:binary`)
  """
  @type source :: :path | :binary

  @typedoc """
  Option to specify the XLSX file source
  """
  @type source_option :: {:source, source()}

  @typedoc """
  List of cell values
  """
  @type row :: list(any())

  @typedoc """
  List of rows
  """
  @type rows :: list(row())

  @typedoc """
  Sheet name
  """
  @type sheet_name :: String.t()

  @typedoc """
  Error tuple with message describing the cause of the error
  """
  @type error :: {:error, String.t()}

  @doc """

  Opens an XLSX file located on the file system (default) or from memory.

  ## Examples

  ### Opening XLSX file on the file system

  ```elixir
  {:ok, package} = XlsxReader.open("test.xlsx")
  ```

  ### Opening XLSX file from memory

  ```elixir
  blob = File.read!("test.xlsx")

  {:ok, package} = XlsxReader.open(blob, source: :binary)
  ```

  ## Options

    * `source`: `:path` (on the file system, default) or `:binary` (in memory)
    * `supported_custom_formats`: a list of `{regex | string, type}` tuples (see "Additional custom formats support")
    * `exclude_hidden_sheets?`: Whether to exclude hidden sheets in the workbook

  """
  @spec open(String.t() | binary(), [source_option]) ::
          {:ok, XlsxReader.Package.t()} | error()
  def open(file, options \\ []) do
    file
    |> ZipArchive.handle(Keyword.get(options, :source, :path))
    |> PackageLoader.open(
      Keyword.take(options, [:supported_custom_formats, :exclude_hidden_sheets?])
    )
  end

  @doc """

  Lists the names of the sheets in the package's workbook

  """
  @spec sheet_names(XlsxReader.Package.t()) :: [sheet_name()]
  def sheet_names(package) do
    for %{name: name} <- package.workbook.sheets, do: name
  end

  @doc """

  Loads the sheet with the given name (see `sheet_names/1`)

  ## Options

    * `type_conversion` - boolean (default: `true`)
    * `blank_value` - placeholder value for empty cells (default: `""`)
    * `empty_rows` - include empty rows (default: `true`)
    * `number_type` - type used for numeric conversion :`Integer`, `Decimal` or `Float` (default: `Float`)
    * `skip_row?`: function callback that determines if a row should be skipped.
       Takes precedence over `blank_value` and `empty_rows`.
       Defaults to `nil` (keeping the behaviour of `blank_value` and `empty_rows`).
    * `cell_data_format`: Controls the format of the cell data. Can be `:value` (default, returns the cell value only) or `:cell` (returns instances of `XlsxReader.Cell`).

  The `Decimal` type requires the [decimal](https://github.com/ericmj/decimal) library.

  ## Examples

  ### Skipping rows

  When using the `skip_row?` callback, rows are ignored in the parser which is more memory efficient.

  ```elixir
  # Skip all rows for which all the values are either blank or "-"
  XlsxReader.sheet(package, "Sheet1", skip_row?: fn row ->
    Enum.all?(row, & String.trim(&1) in ["", "-"])
  end)

  # Skip all rows for which the first column contains the text "disabled"
  XlsxReader.sheet(package, "Sheet1", skip_row?: fn [column | _] ->
    column == "disabled"
  end)
  ```

  """
  @spec sheet(XlsxReader.Package.t(), sheet_name(), Keyword.t()) :: {:ok, rows()} | error()
  def sheet(package, sheet_name, options \\ []) do
    PackageLoader.load_sheet_by_name(package, sheet_name, options)
  end

  @doc """

  Loads all the sheets in the workbook.

  On success, returns `{:ok, [{sheet_name, rows}, ...]}`.

  ## Filtering options

    * `only` - include the sheets whose name matches the filter
    * `except` - exclude the sheets whose name matches the filter

  Sheets can filtered by name using:

    * a string (e.g. `"Exact Match"`)
    * a regex (e.g. `~r/Sheet \d+/`)
    * a list of string and/or regexes (e.g. `["Parameters", ~r/Sheet [12]/]`)

  ## Sheet options

  See `sheet/2`.

  """
  @spec sheets(XlsxReader.Package.t(), Keyword.t()) ::
          {:ok, list({sheet_name(), rows()})} | error()
  def sheets(package, options \\ []) do
    package.workbook.sheets
    |> filter_sheets_by_name(
      sheet_filter_option(options, :only),
      sheet_filter_option(options, :except)
    )
    |> Enum.reduce_while([], fn sheet, acc ->
      case PackageLoader.load_sheet_by_rid(package, sheet.rid, options) do
        {:ok, rows} ->
          {:cont, [{sheet.name, rows} | acc]}

        error ->
          {:halt, error}
      end
    end)
    |> case do
      sheets when is_list(sheets) ->
        {:ok, Enum.reverse(sheets)}

      error ->
        error
    end
  end

  @doc """

  Loads all the sheets in the workbook concurrently.

  On success, returns `{:ok, [{sheet_name, rows}, ...]}`.

  When processing files with multiple sheets, `async_sheets/3` is ~3x faster than `sheets/2`
  but it comes with a caveat. `async_sheets/3` uses `Task.async_stream/3` under the hood and thus
  runs each concurrent task with a timeout. If you expect your dataset to be of a significant size,
  you may want to increase it from the default 10000ms (see "Concurrency options" below).

  If the order in which the sheets are returned is not relevant for your application, you can
  pass `ordered: false` (see "Concurrency options" below) for a modest speed gain.

  ## Filtering options

  See `sheets/2`.

  ## Sheet options

  See `sheet/2`.

  ## Concurrency options

    * `max_concurrency` - maximum number of tasks to run at the same time (default: `System.schedulers_online/0`)
    * `ordered` - maintain order consistent with `sheet_names/1` (default: `true`)
    * `timeout` - maximum duration in milliseconds to process a sheet (default: `10_000`)

  """
  def async_sheets(package, sheet_options \\ [], task_options \\ []) do
    max_concurrency = Keyword.get(task_options, :max_concurrency, System.schedulers_online())
    ordered = Keyword.get(task_options, :ordered, true)
    timeout = Keyword.get(task_options, :timeout, 10_000)

    package.workbook.sheets
    |> filter_sheets_by_name(
      sheet_filter_option(sheet_options, :only),
      sheet_filter_option(sheet_options, :except)
    )
    |> Task.async_stream(
      fn sheet ->
        case PackageLoader.load_sheet_by_rid(package, sheet.rid, sheet_options) do
          {:ok, rows} ->
            {:ok, {sheet.name, rows}}

          error ->
            error
        end
      end,
      max_concurrency: max_concurrency,
      ordered: ordered,
      timeout: timeout,
      on_timeout: :kill_task
    )
    |> Enum.reduce_while({:ok, []}, fn
      {:ok, {:ok, entry}}, {:ok, acc} ->
        {:cont, {:ok, [entry | acc]}}

      {:ok, error}, _acc ->
        {:halt, {:error, error}}

      {:exit, :timeout}, _acc ->
        {:halt, {:error, "timeout exceeded"}}

      {:exit, reason}, _acc ->
        {:halt, {:error, reason}}
    end)
    |> case do
      {:ok, list} ->
        if ordered,
          do: {:ok, Enum.reverse(list)},
          else: {:ok, list}

      error ->
        error
    end
  end

  ## Sheet filter

  def sheet_filter_option(options, key),
    do: options |> Keyword.get(key, []) |> List.wrap()

  defp filter_sheets_by_name(sheets, [], []), do: sheets

  defp filter_sheets_by_name(sheets, only, except) do
    Enum.filter(sheets, fn %{name: name} ->
      filter_only?(name, only) && !filter_except?(name, except)
    end)
  end

  defp filter_only?(_name, []), do: true
  defp filter_only?(name, filters), do: Enum.any?(filters, &filter_match?(name, &1))

  defp filter_except?(_name, []), do: false
  defp filter_except?(name, filters), do: Enum.any?(filters, &filter_match?(name, &1))

  defp filter_match?(name, %Regex{} = regex), do: String.match?(name, regex)
  defp filter_match?(exact_match, exact_match) when is_binary(exact_match), do: true
  defp filter_match?(_, _), do: false
end

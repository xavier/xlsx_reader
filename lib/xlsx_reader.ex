defmodule XlsxReader do
  @moduledoc """

  Opens XLSX workbook and reads its worksheets.

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

  """
  @spec open(String.t() | binary(), [source_option]) ::
          {:ok, XlsxReader.Package.t()} | error()
  def open(file, options \\ []) do
    file
    |> ZipArchive.handle(Keyword.get(options, :source, :path))
    |> PackageLoader.open()
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
    * `number_type` - type used for numeric conversion :`Integer`, 'Decimal' or `Float` (default: `Float`)

  The `Decimal` type requires the [decimal](https://github.com/ericmj/decimal) library.

  """
  @spec sheet(XlsxReader.Package.t(), sheet_name(), Keyword.t()) :: {:ok, rows()}
  def sheet(package, sheet_name, options \\ []) do
    PackageLoader.load_sheet_by_name(package, sheet_name, options)
  end

  @doc """

  Loads all the sheets in the workbook.

  On success, returns a list of `{sheet_name, rows}` pairs.

  ## Options

  See `sheet/2`.

  """
  @spec sheets(XlsxReader.Package.t(), Keyword.t()) ::
          {:ok, list({sheet_name(), rows()})} | error()
  def sheets(package, options \\ []) do
    load_all_sheets(package, options, package.workbook.sheets, [])
  end

  ##

  @spec load_all_sheets(
          XlsxReader.Package.t(),
          Keyword.t(),
          list(XlsxReader.Sheet.t()),
          list({sheet_name(), rows()})
        ) ::
          {:ok, list({sheet_name(), rows()})} | error()

  defp load_all_sheets(_package, _options, [], acc), do: {:ok, Enum.reverse(acc)}

  defp load_all_sheets(package, options, [sheet | sheets], acc) do
    case PackageLoader.load_sheet_by_rid(package, sheet.rid, options) do
      {:ok, rows} ->
        load_all_sheets(package, options, sheets, [{sheet.name, rows} | acc])

      error ->
        error
    end
  end
end

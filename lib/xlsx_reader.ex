defmodule XlsxReader do
  @moduledoc """

  Opens XLSX workbook and reads its worksheets

  """

  alias XlsxReader.{Package, Unzip}

  @type source :: :binary | :path
  @type source_option :: {:source, source()}
  @type row :: list(any())
  @type rows :: list(row())
  @type error :: {:error, String.t()}

  @doc """

  Opens an XLSX file from the file system (default) or from memory.

  ## Options

    * `source`: `:path` (on the file system, default) or `:binary` (in memory)

  """
  @spec open(String.t() | binary(), [source_option]) ::
          {:ok, XlsxReader.Package.t()} | error()
  def open(source, options \\ []) do
    source
    |> Unzip.handle(Keyword.get(options, :source, :path))
    |> Package.open()
  end

  @doc """

  Returns the names of the sheets in the package's workbook

  """
  @spec sheet_names(XlsxReader.Package.t()) :: [String.t()]
  def sheet_names(package) do
    for %{name: name} <- package.workbook.sheets, do: name
  end

  @doc """

  Parses the sheet with the given name (see `sheet_names/1`)

  ## Options

    * `type_conversion` - boolean (default: `true`)
    * `blank_value` - placeholder value for empty cells (default: `""`)
    * `empty_rows` - include empty rows (default: `true`)
    * `number_type` - type used for numeric conversion :`Integer`, 'Decimal' or `Float` (default: `Float`)

  The `Decimal` type requires the [decimal](https://github.com/ericmj/decimal) library.

  """
  @spec sheet(XlsxReader.Package.t(), String.t(), Keyword.t()) :: {:ok, rows()}
  def sheet(package, sheet_name, options \\ []) do
    Package.load_sheet_by_name(package, sheet_name, options)
  end

  @doc """

  Parses all the sheets in the workbook.

  ## Options

  See `sheet/2`.

  """
  @spec sheets(XlsxReader.Package.t(), Keyword.t()) :: {:ok, rows()}
  def sheets(package, options \\ []) do
    {:ok, Package.load_sheets(package, options)}
  end
end

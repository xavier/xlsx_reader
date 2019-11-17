defmodule XlsxReader do
  @moduledoc """

  Opens XLSX workbook and reads its worksheets

  """

  alias XlsxReader.{Package, Unzip}

  @type source :: :binary | :path
  @type option :: {:source, source()}
  @type row :: list(any())
  @type rows :: list(row())

  @doc """

  Opens an XLSX file

  Options:

  - `source`: `:binary` or `:path`

  """
  @spec open(String.t() | binary(), [option]) ::
          {:ok, XlsxReader.Package.t()} | {:error, String.t()}
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

  """
  @spec sheet(XlsxReader.Package.t(), String.t()) :: {:ok, rows()}
  def sheet(package, sheet_name) do
    Package.load_sheet_by_name(package, sheet_name)
  end

  @doc """

  Parses all the sheets in the workbook.

  """
  @spec sheets(XlsxReader.Package.t()) :: {:ok, rows()}
  def sheets(package) do
    {:ok, Package.load_sheets(package)}
  end
end

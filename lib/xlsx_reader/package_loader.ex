defmodule XlsxReader.PackageLoader do
  @moduledoc false

  # Loads the content of an XLSX file.
  #
  # An XLSX file is ZIP archive containing XML files linked to each other
  # using relationships defined in `_rels/*.xml.rels` files.

  alias XlsxReader.Parsers.{
    RelationshipsParser,
    SharedStringsParser,
    StylesParser,
    WorkbookParser,
    WorksheetParser
  }

  alias XlsxReader.ZipArchive

  @doc """

  Opens an XLSX package.

  It verifies the contents of the archive and preloads the workbook sheet list
  and relationships as well as the shared strings and style information required
  to load the sheet data.

  To load the actual sheet data, see `load_sheet_by_rid/3` and `load_sheet_by_name/3`.

  """
  @spec open(XlsxReader.ZipArchive.zip_handle()) ::
          {:ok, XlsxReader.Package.t()} | XlsxReader.error()
  def open(zip_handle, options \\ []) do
    with :ok <- check_contents(zip_handle),
         {:ok, workbook} <- load_workbook_xml(zip_handle),
         {:ok, workbook_rels} <- load_workbook_xml_rels(zip_handle) do
      package =
        %XlsxReader.Package{
          zip_handle: zip_handle,
          workbook: %{workbook | rels: workbook_rels}
        }
        |> load_shared_strings
        |> load_styles(Keyword.get(options, :supported_custom_formats, []))

      {:ok, package}
    end
  end

  @doc """
  Loads a single sheet identified by relationship id (`rId`)

  ## Options

  See `XlsxReader.sheet/2`.

  """
  @spec load_sheet_by_rid(XlsxReader.Package.t(), String.t(), Keyword.t()) ::
          {:ok, XlsxReader.row()} | XlsxReader.error()
  def load_sheet_by_rid(package, rid, options \\ []) do
    case fetch_rel_target(package.workbook.rels, :sheets, rid) do
      {:ok, target} ->
        load_worksheet_xml(package, xl_path(target), options)

      :error ->
        {:error, "sheet relationship not found"}
    end
  end

  @doc """
  Loads a single sheet identified by name

  ## Options

  See `XlsxReader.sheet/2`.

  """
  @spec load_sheet_by_name(XlsxReader.Package.t(), String.t(), Keyword.t()) ::
          {:ok, XlsxReader.row()} | XlsxReader.error()
  def load_sheet_by_name(package, name, options \\ []) do
    case find_sheet_by_name(package, name) do
      %{rid: rid} ->
        load_sheet_by_rid(package, rid, options)

      nil ->
        {:error, "sheet #{inspect(name)} not found"}
    end
  end

  ##

  @workbook_xml "xl/workbook.xml"
  @workbook_xml_rels "xl/_rels/workbook.xml.rels"

  @required_files [
    "[Content_Types].xml",
    @workbook_xml,
    @workbook_xml_rels
  ]

  defp check_contents(zip_handle) do
    with {:ok, files} <- ZipArchive.list(zip_handle) do
      if Enum.all?(@required_files, &Enum.member?(files, &1)),
        do: :ok,
        else: {:error, "invalid xlsx file"}
    end
  end

  defp load_workbook_xml(zip_handle) do
    with {:ok, xml} <- ZipArchive.extract(zip_handle, @workbook_xml) do
      WorkbookParser.parse(xml)
    end
  end

  defp load_workbook_xml_rels(zip_handle) do
    with {:ok, xml} <- ZipArchive.extract(zip_handle, @workbook_xml_rels) do
      RelationshipsParser.parse(xml)
    end
  end

  defp load_shared_strings(package) do
    with {:ok, file} <- single_rel_target(package.workbook.rels.shared_strings),
         {:ok, xml} <- ZipArchive.extract(package.zip_handle, file),
         {:ok, shared_strings} <- SharedStringsParser.parse(xml) do
      %{package | workbook: %{package.workbook | shared_strings: shared_strings}}
    end
  end

  defp load_styles(package, supported_custom_formats) do
    with {:ok, file} <- single_rel_target(package.workbook.rels.styles),
         {:ok, xml} <- ZipArchive.extract(package.zip_handle, file),
         {:ok, style_types, custom_formats} <- StylesParser.parse(xml, supported_custom_formats) do
      %{
        package
        | workbook: %{package.workbook | style_types: style_types, custom_formats: custom_formats}
      }
    end
  end

  defp single_rel_target(rels) do
    case Map.values(rels) do
      [target] ->
        {:ok, xl_path(target)}

      targets ->
        {:error, "expected a single target in #{inspect(targets)}"}
    end
  end

  defp load_worksheet_xml(package, file, options) do
    with {:ok, xml} <- ZipArchive.extract(package.zip_handle, file) do
      WorksheetParser.parse(xml, package.workbook, options)
    end
  end

  defp xl_path(relative_path), do: Path.join("xl", relative_path)

  defp find_sheet_by_name(package, name) do
    Enum.find(package.workbook.sheets, fn %{name: n} -> name == n end)
  end

  defp fetch_rel_target(rels, type, rid) do
    with {:ok, paths} <- Map.fetch(rels, type) do
      Map.fetch(paths, rid)
    end
  end
end

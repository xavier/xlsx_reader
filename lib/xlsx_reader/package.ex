defmodule XlsxReader.Package do
  @enforce_keys [:zip_handle]
  defstruct zip_handle: nil, files: [], workbook: nil

  alias XlsxReader.{
    Unzip,
    RelationshipsParser,
    SharedStringsParser,
    WorkbookParser,
    WorksheetParser
  }

  def open(zip_handle) do
    with {:ok, files} <- list_contents(zip_handle),
         {:ok, sheets} <- load_workbook_xml(zip_handle),
         {:ok, workbook_rels} <- load_workbook_xml_rels(zip_handle) do
      package =
        %__MODULE__{
          zip_handle: zip_handle,
          files: files,
          workbook: %{
            sheets: sheets,
            rels: workbook_rels,
            shared_strings: nil
          }
        }
        |> load_shared_strings

      {:ok, package}
    end
  end

  def load_sheets(package) do
    for sheet <- package.workbook.sheets do
      with {:ok, target} <- fetch_rel_target(package.workbook.rels, :sheets, sheet.rid),
           file <- xl_path(target),
           {:ok, rows} <- load_worksheet_xml(package, file) do
        {sheet.name, rows}
      else
        :error ->
          {:error, "sheet relationship not found"}
      end
    end
  end

  def load_sheet(package, name) do
    with %{rid: rid} <- find_sheet_by_name(package, name),
         {:ok, target} <- fetch_rel_target(package.workbook.rels, :sheets, rid),
         file <- xl_path(target) do
      load_worksheet_xml(package, file)
    else
      nil ->
        {:error, "sheet #{inspect(name)} not found"}

      :error ->
        {:error, "sheet relationship not found"}
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

  def list_contents(zip_handle) do
    with {:ok, files} <- Unzip.list(zip_handle) do
      if Enum.all?(@required_files, &Enum.member?(files, &1)) do
        {:ok, files}
      else
        {:error, "invalid xlsx file"}
      end
    end
  end

  defp load_workbook_xml(zip_handle) do
    with {:ok, xml} <- Unzip.extract(zip_handle, @workbook_xml) do
      WorkbookParser.parse(xml)
    end
  end

  defp load_workbook_xml_rels(zip_handle) do
    with {:ok, xml} <- Unzip.extract(zip_handle, @workbook_xml_rels) do
      RelationshipsParser.parse(xml)
    end
  end

  defp load_shared_strings(package) do
    with {:ok, file} <- shared_strings_file(package),
         {:ok, xml} <- Unzip.extract(package.zip_handle, file),
         {:ok, shared_strings} <- SharedStringsParser.parse(xml) do
      %{package | workbook: Map.put(package.workbook, :shared_strings, shared_strings)}
    else
      :no_shared_strings ->
        package
    end
  end

  defp shared_strings_file(package) do
    case map_size(package.workbook.rels.shared_strings) do
      1 ->
        path =
          package.workbook.rels.shared_strings
          |> Map.values()
          |> List.first()
          |> xl_path()

        {:ok, path}

      0 ->
        :no_shared_strings

      _ ->
        {:error, "more than one sharedString relationship"}
    end
  end

  def load_worksheet_xml(package, file) do
    with {:ok, xml} <- Unzip.extract(package.zip_handle, file) do
      WorksheetParser.parse(xml, package.workbook.shared_strings)
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

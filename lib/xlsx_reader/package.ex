defmodule XlsxReader.Package do
  @moduledoc """

  Loads the content of an XLSX file.

  """

  @enforce_keys [:zip_handle, :workbook]
  defstruct zip_handle: nil, workbook: nil

  @type t :: %__MODULE__{
          zip_handle: XlsxReader.Unzip.zip_handle(),
          workbook: XlsxReader.Workbook.t()
        }

  alias XlsxReader.{
    RelationshipsParser,
    SharedStringsParser,
    StylesParser,
    Unzip,
    WorkbookParser,
    WorksheetParser
  }

  def open(zip_handle) do
    with :ok <- check_contents(zip_handle),
         {:ok, workbook} <- load_workbook_xml(zip_handle),
         {:ok, workbook_rels} <- load_workbook_xml_rels(zip_handle) do
      package =
        %__MODULE__{
          zip_handle: zip_handle,
          workbook: %{workbook | rels: workbook_rels}
        }
        |> load_shared_strings
        |> load_styles

      {:ok, package}
    end
  end

  def load_sheets(package) do
    for sheet <- package.workbook.sheets do
      case load_sheet_by_rid(package, sheet.rid) do
        {:ok, rows} ->
          {sheet.name, rows}

        _ ->
          {sheet.name, []}
      end
    end
  end

  def load_sheet_by_rid(package, rid) do
    case fetch_rel_target(package.workbook.rels, :sheets, rid) do
      {:ok, target} ->
        load_worksheet_xml(package, xl_path(target))

      :error ->
        {:error, "sheet relationship not found"}
    end
  end

  def load_sheet_by_name(package, name) do
    case find_sheet_by_name(package, name) do
      %{rid: rid} ->
        load_sheet_by_rid(package, rid)

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

  def check_contents(zip_handle) do
    with {:ok, files} <- Unzip.list(zip_handle) do
      if Enum.all?(@required_files, &Enum.member?(files, &1)),
        do: :ok,
        else: {:error, "invalid xlsx file"}
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
    with {:ok, file} <- single_rel_target(package.workbook.rels.shared_strings),
         {:ok, xml} <- Unzip.extract(package.zip_handle, file),
         {:ok, shared_strings} <- SharedStringsParser.parse(xml) do
      %{package | workbook: Map.put(package.workbook, :shared_strings, shared_strings)}
    else
      :no_shared_strings ->
        package
    end
  end

  defp load_styles(package) do
    with {:ok, file} <- single_rel_target(package.workbook.rels.styles),
         {:ok, xml} <- Unzip.extract(package.zip_handle, file),
         {:ok, style_types} <- StylesParser.parse(xml) do
      %{package | workbook: %{package.workbook | style_types: style_types}}
    else
      :no_shared_strings ->
        package
    end
  end

  defp single_rel_target(rels) do
    case Map.values(rels) do
      [target] ->
        {:ok, xl_path(target)}

      [] ->
        :no_shared_strings

      _ ->
        {:error, "more than one sharedString relationship"}
    end
  end

  def load_worksheet_xml(package, file) do
    with {:ok, xml} <- Unzip.extract(package.zip_handle, file) do
      WorksheetParser.parse(xml, package.workbook)
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

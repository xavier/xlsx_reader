defmodule XlsxReader.Parsers.WorkbookParser do
  @moduledoc false

  # Parses SpreadsheetML workbooks.
  #
  # The `workbook.xml` contains the worksheet list and their relationship identifier (`rId`).
  # The workbook may also contain a hint regarding the [date system](https://docs.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system) in use.
  #

  @behaviour Saxy.Handler

  alias XlsxReader.Conversion
  alias XlsxReader.Parsers.Utils

  @doc """
  Parses a workbook XML document.

  ## Options
    * `:exclude_hidden_sheets?` - Whether to exclude hidden sheets in the workbook
    * `:exclude_empty_sheets?` - Whether to exclude empty sheets in the workbook
    * `:preload_sheets?` - Whether to preload all workbook sheets on open
  """
  def parse(xml, options \\ []) do
    exclude_empty_sheets? = Keyword.get(options, :exclude_empty_sheets?, false)
    exclude_hidden_sheets? = Keyword.get(options, :exclude_hidden_sheets?, false)
    preload_sheets? = exclude_empty_sheets? or Keyword.get(options, :preload_sheets?, false)

    Saxy.parse_string(xml, __MODULE__, %XlsxReader.Workbook{
      options: %{
        exclude_empty_sheets?: exclude_empty_sheets?,
        exclude_hidden_sheets?: exclude_hidden_sheets?,
        preload_sheets?: preload_sheets?
      }
    })
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, workbook) do
    {:ok, workbook}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, workbook) do
    {:ok, %{workbook | base_date: workbook.base_date || Conversion.base_date(1900)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"workbookPr", attributes}, workbook) do
    {:ok, %{workbook | base_date: attributes |> date_system() |> Conversion.base_date()}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"sheet", attributes}, workbook) do
    is_hidden? = attributes |> Utils.get_attribute("state") === "hidden"
    skip_sheet? = workbook.options.exclude_hidden_sheets? && is_hidden?

    case skip_sheet? do
      true -> {:ok, workbook}
      false -> {:ok, %{workbook | sheets: [build_sheet(attributes) | workbook.sheets]}}
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, workbook) do
    {:ok, workbook}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sheets", workbook) do
    {:ok, %{workbook | sheets: Enum.reverse(workbook.sheets)}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, workbook) do
    {:ok, workbook}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, workbook) do
    {:ok, workbook}
  end

  ##

  @sheet_attributes %{
    "name" => :name,
    "r:id" => :rid,
    "sheetId" => :sheet_id
  }

  defp build_sheet(attributes) do
    Utils.map_attributes(attributes, @sheet_attributes, %XlsxReader.Sheet{})
  end

  defp date_system(attributes) do
    if Utils.get_attribute(attributes, "date1904", "0") == "1",
      do: 1904,
      else: 1900
  end
end

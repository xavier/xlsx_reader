defmodule XlsxReader.Parsers.WorkbookParser do
  @moduledoc false

  # Parses SpreadsheetML workbooks.
  #
  # The `workbook.xml` contains the worksheet list and their relationship identifier (`rId`).
  # The workbook may also contain a hint regarding the [date system](https://docs.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system) in use.
  #

  @behaviour Saxy.Handler

  alias XlsxReader.{Conversion, Workbook}
  alias XlsxReader.Parsers.Utils

  @doc """
  Parses a workbook XML document.

  ## Options
    * `:exclude_hidden_sheets?` - Whether to exclude hidden sheets in the workbook
  """
  def parse(xml, options \\ []) do
    exclude_hidden_sheets? = Keyword.get(options, :exclude_hidden_sheets?, false)

    Saxy.parse_string(xml, __MODULE__, %Workbook{
      options: %{exclude_hidden_sheets?: exclude_hidden_sheets?}
    })
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, %Workbook{} = workbook) do
    {:ok, workbook}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, %Workbook{} = workbook) do
    {:ok, %{workbook | base_date: workbook.base_date || Conversion.base_date(1900)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"workbookPr", attributes}, %Workbook{} = workbook) do
    {:ok, %{workbook | base_date: attributes |> date_system() |> Conversion.base_date()}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"sheet", attributes}, %Workbook{} = workbook) do
    is_hidden? = attributes |> Utils.get_attribute("state") === "hidden"
    skip_sheet? = workbook.options.exclude_hidden_sheets? && is_hidden?

    case skip_sheet? do
      true -> {:ok, workbook}
      false -> {:ok, %{workbook | sheets: [build_sheet(attributes) | workbook.sheets]}}
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, %Workbook{} = workbook) do
    {:ok, workbook}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sheets", %Workbook{} = workbook) do
    {:ok, %{workbook | sheets: Enum.reverse(workbook.sheets)}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, %Workbook{} = workbook) do
    {:ok, workbook}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, %Workbook{} = workbook) do
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

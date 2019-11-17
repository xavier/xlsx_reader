defmodule XlsxReader.WorkbookParser do
  @behaviour Saxy.Handler

  def parse(xml) do
    Saxy.parse_string(xml, __MODULE__, %XlsxReader.Workbook{})
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"sheet", attributes}, state) do
    {:ok, %{state | sheets: [build_sheet(attributes) | state.sheets]}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sheets", state) do
    {:ok, %{state | sheets: Enum.reverse(state.sheets)}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, state) do
    {:ok, state}
  end

  ##

  @sheet_attributes %{
    "name" => :name,
    "r:id" => :rid,
    "sheetId" => :sheet_id
  }

  defp build_sheet(attributes) do
    Enum.reduce(
      attributes,
      %XlsxReader.Sheet{},
      fn {name, value}, sheet ->
        case Map.fetch(@sheet_attributes, name) do
          {:ok, key} ->
            %{sheet | key => value}

          :error ->
            sheet
        end
      end
    )
  end
end

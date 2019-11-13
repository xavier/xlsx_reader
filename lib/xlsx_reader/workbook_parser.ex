defmodule XlsxReader.WorkbookParser do
  @behaviour Saxy.Handler

  def parse(xml) do
    Saxy.parse_string(xml, __MODULE__, [])
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
    {:ok, [map_sheet_attributes(attributes) | state]}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sheets", state) do
    {:ok, Enum.reverse(state)}
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

  @attributes_mapping %{
    "name" => :name,
    "r:id" => :rid,
    "sheetId" => :sheet_id
  }

  defp map_sheet_attributes(attributes) do
    Enum.reduce(attributes, %{}, fn {name, value}, acc ->
      case Map.fetch(@attributes_mapping, name) do
        {:ok, key} ->
          Map.put(acc, key, value)

        :error ->
          acc
      end
    end)
  end
end

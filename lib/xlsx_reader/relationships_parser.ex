defmodule XlsxReader.RelationshipsParser do
  @behaviour Saxy.Handler

  def parse(xml) do
    Saxy.parse_string(xml, __MODULE__, %{
      shared_strings: %{},
      styles: %{},
      themes: %{},
      sheets: %{}
    })
  end

  @namespace "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  @types %{
    "#{@namespace}/sharedStrings" => :shared_strings,
    "#{@namespace}/styles" => :styles,
    "#{@namespace}/theme" => :themes,
    "#{@namespace}/worksheet" => :sheets
  }

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"Relationship", attributes}, state) do
    with %{id: id, target: target, type: type} <- extract_relationship_attributes(attributes),
         {:ok, key} <- Map.fetch(@types, type) do
      {:ok, Map.update!(state, key, fn rels -> Map.put_new(rels, id, target) end)}
    else
      _ ->
        {:ok, state}
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
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
    "Id" => :id,
    "Target" => :target,
    "Type" => :type
  }

  defp extract_relationship_attributes(attributes) do
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

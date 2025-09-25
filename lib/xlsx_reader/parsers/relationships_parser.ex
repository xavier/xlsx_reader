defmodule XlsxReader.Parsers.RelationshipsParser do
  @moduledoc false

  # Parses SpreadsheetML workbook relationships.
  #
  # The relationships determine the exact location of the shared strings, styles,
  # themes and worksheet files within the archive.

  alias XlsxReader.Parsers.Utils

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
      {:ok,
       Map.update!(state, key, fn rels -> Map.put_new(rels, id, sanitize_target(target)) end)}
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

  @relationship_attributes_mapping %{
    "Id" => :id,
    "Target" => :target,
    "Type" => :type
  }

  defp extract_relationship_attributes(attributes) do
    Utils.map_attributes(attributes, @relationship_attributes_mapping)
  end

  # Remove leading "/" to deal with file containing absolute paths
  defp sanitize_target("/" <> target), do: target
  defp sanitize_target(target), do: target
end

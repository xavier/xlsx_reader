defmodule XlsxReader.SharedStringsParser do
  @moduledoc """

  Builds a list of shared strings referenced to by index in the worksheets

  """

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
  def handle_event(:start_element, {"t", _attributes}, state) do
    {:ok, [:expect_chars | state]}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "sst", state) do
    {:ok, Enum.reverse(state)}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, chars, [:expect_chars | state]) do
    {:ok, [chars | state]}
  end
end

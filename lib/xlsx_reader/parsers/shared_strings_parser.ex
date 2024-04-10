defmodule XlsxReader.Parsers.SharedStringsParser do
  @moduledoc false

  # Parses SpreadsheetML shared strings definitions.
  #
  # The parser builds a list of shared strings.
  #
  # Worksheets only contain numbers: the value of cells containing text is
  # numeric index to the array of shared strings.

  @behaviour Saxy.Handler

  alias XlsxReader.Array
  alias XlsxReader.Parsers.Utils

  defmodule State do
    @moduledoc false

    defstruct current_string: nil,
              strings: [],
              expect_chars: false,
              preserve_space: false
  end

  def parse(xml) do
    xml
    |> Utils.strip_leading_bom()
    |> Saxy.parse_string(__MODULE__, %State{})
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, Array.from_list(Enum.reverse(state.strings))}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"si", _attributes}, state) do
    {:ok, %{state | current_string: ""}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"t", attributes}, state) do
    {:ok,
     %{
       state
       | expect_chars: true,
         preserve_space: Utils.get_attribute(attributes, "xml:space") == "preserve"
     }}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "t", state) do
    {:ok, %{state | expect_chars: false, preserve_space: false}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "si", state) do
    {:ok, %{state | current_string: nil, strings: [state.current_string | state.strings]}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, chars, %{expect_chars: true} = state) do
    {:ok, %{state | current_string: state.current_string <> preserve_space(state, chars)}}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, %{expect_chars: false} = state) do
    {:ok, state}
  end

  ##
  defp preserve_space(%{preserve_space: true}, string), do: string
  defp preserve_space(%{preserve_space: false}, string), do: String.trim(string)
end

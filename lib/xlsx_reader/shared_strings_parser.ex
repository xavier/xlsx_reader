defmodule XlsxReader.SharedStringsParser do
  @moduledoc false

  # Parses SpreadsheetML shared strings definitions.
  #
  # The parser builds a list of shared strings.
  #
  # Worksheets only contain numbers: the value of cells containing text is
  # numeric index to the array of shared strings.

  @behaviour Saxy.Handler

  alias XlsxReader.ParserUtils

  defmodule State do
    @moduledoc false

    defstruct current_string: nil,
              strings: [],
              expect_chars: false,
              got_chars: false,
              preserve_space: false
  end

  def parse(xml) do
    Saxy.parse_string(xml, __MODULE__, %State{})
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, Enum.reverse(state.strings)}
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
         got_chars: false,
         preserve_space: ParserUtils.get_attribute(attributes, "xml:space") == "preserve"
     }}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "t", state) do
    state = preserve_space(state)
    {:ok, %{state | expect_chars: false, got_chars: false, preserve_space: false}}
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
    {:ok, %{state | current_string: state.current_string <> chars, got_chars: true}}
  end

  ##

  # Work around a bug in Saxy: https://github.com/qcam/saxy/issues/51
  # If we expected characters with `xml:space="preserve"` but got nothing, we emit a single space
  defp preserve_space(state) do
    if state.preserve_space && state.got_chars == false do
      %{state | current_string: state.current_string <> " "}
    else
      state
    end
  end
end

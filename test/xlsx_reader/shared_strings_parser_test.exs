defmodule XlsxReader.SharedStringsParserTest do
  use ExUnit.Case

  alias XlsxReader.SharedStringsParser

  test "parses sharedStrings.xml" do
    shared_strings_xml = TestFixtures.read!("package/xl/sharedStrings.xml")

    expected = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "some ",
      "test",
      "date",
      "datetime",
      "time",
      "percentage",
      "money chf",
      "money usd",
      "9999,99 USD",
      "ticked",
      "not ticked",
      "hyperlink",
      "https://elixir-lang.org/"
    ]

    assert {:ok, expected} == SharedStringsParser.parse(shared_strings_xml)
  end

  test "parses strings with rich text" do
    shared_strings_xml = TestFixtures.read!("xml/sharedStringsWithRichText.xml")

    expected = [
      "Cell A1",
      "Cell B1",
      "My Cell",
      "Cell A2",
      "Cell B2"
    ]

    assert {:ok, expected} == SharedStringsParser.parse(shared_strings_xml)
  end
end

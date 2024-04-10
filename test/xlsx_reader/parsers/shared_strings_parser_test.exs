defmodule XlsxReader.Parsers.SharedStringsParserTest do
  use ExUnit.Case

  alias XlsxReader.Array
  alias XlsxReader.Parsers.SharedStringsParser

  test "parses sharedStrings.xml" do
    shared_strings_xml = TestFixtures.read!("package/xl/sharedStrings.xml")

    expected =
      Array.from_list([
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
        "https://elixir-lang.org/",
        " ",
        "-"
      ])

    assert {:ok, expected} == SharedStringsParser.parse(shared_strings_xml)
  end

  test "parses strings with rich text" do
    shared_strings_xml = TestFixtures.read!("xml/sharedStringsWithRichText.xml")

    expected =
      Array.from_list([
        "Cell A1",
        "Cell B1",
        "My Cell",
        "Cell A2",
        "Cell B2"
      ])

    assert {:ok, expected} == SharedStringsParser.parse(shared_strings_xml)
  end

  test "parses strings with leading BOM" do
    shared_strings_xml = TestFixtures.read!("xml/sharedStringsWithRichText.xml")
    assert {:ok, _} = SharedStringsParser.parse("\uFEFF" <> shared_strings_xml)
  end

  test "takes xml:space instruction into account" do
    shared_strings_xml = TestFixtures.read!("xml/sharedStringsWithXmlSpacePreserve.xml")

    expected =
      Array.from_list([
        "  with spaces  ",
        "without spaces",
        "without spaces"
      ])

    assert {:ok, expected} == SharedStringsParser.parse(shared_strings_xml)
  end
end

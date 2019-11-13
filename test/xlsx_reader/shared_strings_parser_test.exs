defmodule XlsxReader.SharedStringsParserTest do
  use ExUnit.Case

  alias XlsxReader.SharedStringsParser

  test "parses sharedStrings.xml" do
    shared_strings_xml = TestFixtures.read!("package/xl/sharedStrings.xml")

    expected = [
      "Table 1",
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "some ",
      "test"
    ]

    assert {:ok, expected} == SharedStringsParser.parse(shared_strings_xml)
  end
end

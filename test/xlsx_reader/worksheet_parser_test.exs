defmodule XlsxReader.WorksheetParserTest do
  use ExUnit.Case

  alias XlsxReader.WorksheetParser

  test "parses sheet.xml" do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet1.xml")

    shared_strings = [
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

    assert {:ok, result} = WorksheetParser.parse(sheet_xml, shared_strings)
  end
end

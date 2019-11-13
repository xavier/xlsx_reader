defmodule XlsxReader.WorkbookParserTest do
  use ExUnit.Case

  alias XlsxReader.WorkbookParser

  test "parses workbook.xml" do
    workbook_xml = TestFixtures.read!("package/xl/workbook.xml")

    expected = [
      %{name: "Sheet 1", rid: "rId4", sheet_id: "1"},
      %{name: "Sheet 2", rid: "rId5", sheet_id: "2"}
    ]

    assert {:ok, expected} == WorkbookParser.parse(workbook_xml)
  end
end

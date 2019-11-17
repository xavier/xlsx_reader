defmodule XlsxReader.WorkbookParserTest do
  use ExUnit.Case

  alias XlsxReader.WorkbookParser

  test "parses workbook.xml" do
    workbook_xml = TestFixtures.read!("package/xl/workbook.xml")

    expected = %XlsxReader.Workbook{
      sheets: [
        %XlsxReader.Sheet{name: "Sheet 1", rid: "rId4", sheet_id: "1"},
        %XlsxReader.Sheet{name: "Sheet 2", rid: "rId5", sheet_id: "2"},
        %XlsxReader.Sheet{name: "Sheet 3", rid: "rId6", sheet_id: "3"}
      ],
      rels: nil,
      shared_strings: nil,
      style_types: nil,
      base_date: ~D[1899-12-30]
    }

    assert {:ok, expected} == WorkbookParser.parse(workbook_xml)
  end
end

defmodule XlsxReader.RelationshipsParserTest do
  use ExUnit.Case

  alias XlsxReader.RelationshipsParser

  test "parses workbook.xml.rels" do
    workbook_xml_rels = TestFixtures.read!("package/xl/_rels/workbook.xml.rels")

    expected = [
      %{
        id: "rId1",
        target: "sharedStrings.xml",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
      },
      %{
        id: "rId2",
        target: "styles.xml",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
      },
      %{
        id: "rId3",
        target: "theme/theme1.xml",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
      },
      %{
        id: "rId4",
        target: "worksheets/sheet1.xml",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      },
      %{
        id: "rId5",
        target: "worksheets/sheet2.xml",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      }
    ]

    assert {:ok, expected} == RelationshipsParser.parse(workbook_xml_rels)
  end
end

defmodule XlsxReader.Parsers.RelationshipsParserTest do
  use ExUnit.Case

  alias XlsxReader.Parsers.RelationshipsParser

  test "parses workbook.xml.rels" do
    workbook_xml_rels = TestFixtures.read!("package/xl/_rels/workbook.xml.rels")

    expected = %{
      shared_strings: %{
        "rId1" => "sharedStrings.xml"
      },
      styles: %{
        "rId2" => "styles.xml"
      },
      themes: %{
        "rId3" => "theme/theme1.xml"
      },
      sheets: %{
        "rId4" => "worksheets/sheet1.xml",
        "rId5" => "worksheets/sheet2.xml",
        "rId6" => "worksheets/sheet3.xml"
      }
    }

    assert {:ok, expected} == RelationshipsParser.parse(workbook_xml_rels)
  end

  test "parses strings with leading BOM" do
    workbook_xml_rels = TestFixtures.read!("package/xl/_rels/workbook.xml.rels")
    assert {:ok, _} = RelationshipsParser.parse("\uFEFF" <> workbook_xml_rels)
  end
end

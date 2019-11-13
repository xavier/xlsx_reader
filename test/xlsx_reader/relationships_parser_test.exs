defmodule XlsxReader.RelationshipsParserTest do
  use ExUnit.Case

  alias XlsxReader.RelationshipsParser

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
        "rId5" => "worksheets/sheet2.xml"
      }
    }

    assert {:ok, expected} == RelationshipsParser.parse(workbook_xml_rels)
  end
end

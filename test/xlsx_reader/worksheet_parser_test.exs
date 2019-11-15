defmodule XlsxReader.WorksheetParserTest do
  use ExUnit.Case

  alias XlsxReader.{WorksheetParser, SharedStringsParser, StylesParser}

  setup do
    {:ok, shared_strings} =
      "package/xl/sharedStrings.xml"
      |> TestFixtures.read!()
      |> SharedStringsParser.parse()

    {:ok, styles} =
      "package/xl/styles.xml"
      |> TestFixtures.read!()
      |> StylesParser.parse()

    workbook = %{
      shared_strings: shared_strings,
      styles: styles
    }

    {:ok, workbook: workbook}
  end

  test "parses sheet.xml", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet1.xml")

    assert {:ok, result} = WorksheetParser.parse(sheet_xml, workbook)
  end

  test "looks up shared strings", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet1.xml")

    assert {:ok, rows} = WorksheetParser.parse(sheet_xml, workbook)

    assert [["A", "B", "C", "D", "E", "F", "G"] | _] = rows
  end

  test "performs cell type conversions", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet3.xml")

    expected = [
      ["", "", "", "", ""],
      ["date", ~D[2019-11-14], "", "", ""],
      ["datetime", ~N[2019-11-14 18:14:45], "", "", ""],
      ["time", ~N[2019-11-14 18:14:45], "", "", ""],
      ["percentage", 12.5, "", "", ""],
      ["money chf", "100", "", "", ""],
      ["money usd", "9999,99 USD", "", "", ""],
      ["ticked", true, "", "", ""],
      ["not ticked", false, "", "", ""],
      ["", "", "", "", ""],
      ["", "", "", "", ""]
    ]

    assert {:ok, rows} = WorksheetParser.parse(sheet_xml, workbook)

    assert expected == rows
  end
end

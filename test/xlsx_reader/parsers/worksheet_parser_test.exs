defmodule XlsxReader.Parsers.WorksheetParserTest do
  use ExUnit.Case

  alias XlsxReader.{Conversion, Workbook}
  alias XlsxReader.Parsers.{SharedStringsParser, StylesParser, WorksheetParser}

  setup do
    {:ok, shared_strings} =
      "package/xl/sharedStrings.xml"
      |> TestFixtures.read!()
      |> SharedStringsParser.parse()

    {:ok, style_types} =
      "package/xl/styles.xml"
      |> TestFixtures.read!()
      |> StylesParser.parse()

    workbook = %Workbook{
      shared_strings: shared_strings,
      style_types: style_types,
      base_date: Conversion.base_date(1900)
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

  test "performs cell type conversions by default", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet3.xml")

    expected = [
      ["", ""],
      ["date", ~D[2019-11-15]],
      ["datetime", ~N[2019-11-24 11:06:13]],
      ["time", ~N[1904-01-01 18:45:12]],
      ["percentage", 12.5],
      ["money chf", 100],
      ["money usd", "9999,99 USD"],
      ["ticked", true],
      ["not ticked", false],
      ["hyperlink", "https://elixir-lang.org/"]
    ]

    assert {:ok, rows} = WorksheetParser.parse(sheet_xml, workbook)

    assert expected == rows
  end

  test "returns raw values (except shared strings) when type conversion is disabled", %{
    workbook: workbook
  } do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet3.xml")

    expected = [
      ["", ""],
      ["date", "43784"],
      ["datetime", "43793.462650462963"],
      ["time", "1462.781388888889"],
      ["percentage", "0.125"],
      ["money chf", "100"],
      ["money usd", "9999,99 USD"],
      ["ticked", "1"],
      ["not ticked", "0"],
      ["hyperlink", "https://elixir-lang.org/"]
    ]

    assert {:ok, rows} = WorksheetParser.parse(sheet_xml, workbook, type_conversion: false)

    assert expected == rows
  end

  test "handles inline strings", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("xml/worksheetWithInlineStr.xml")

    expected = [["inline string"]]

    assert {:ok, rows} = WorksheetParser.parse(sheet_xml, workbook)

    assert expected == rows
  end
end

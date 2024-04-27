defmodule XlsxReader.Parsers.WorksheetParserTest do
  use ExUnit.Case

  alias XlsxReader.{Cell, Conversion, Workbook}
  alias XlsxReader.Parsers.{SharedStringsParser, StylesParser, WorksheetParser}

  setup do
    {:ok, shared_strings} =
      "package/xl/sharedStrings.xml"
      |> TestFixtures.read!()
      |> SharedStringsParser.parse()

    {:ok, style_types, custom_formats} =
      "package/xl/styles.xml"
      |> TestFixtures.read!()
      |> StylesParser.parse()

    workbook = %Workbook{
      shared_strings: shared_strings,
      style_types: style_types,
      custom_formats: custom_formats,
      base_date: Conversion.base_date(1900)
    }

    {:ok, workbook: workbook}
  end

  test "parses sheet.xml", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet1.xml")

    assert {:ok, _result} = WorksheetParser.parse(sheet_xml, workbook)
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
      ["money chf", "100"],
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

  test "handles correctly empty values", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet5.xml")

    assert {:ok, [["", "", "Hello", "", "0.0"]]} =
             WorksheetParser.parse(sheet_xml, workbook, type_conversion: false)
  end

  test "handles inline strings", %{workbook: workbook} do
    sheet_xml = TestFixtures.read!("xml/worksheetWithInlineStr.xml")

    expected = [["inline string"]]

    assert {:ok, rows} = WorksheetParser.parse(sheet_xml, workbook)

    assert expected == rows
  end

  test "should ignore rows based on skip_row?", %{
    workbook: workbook
  } do
    sheet_xml = TestFixtures.read!("package/xl/worksheets/sheet4.xml")

    ignore_trimmed = fn row -> Enum.all?(row, &(String.trim(&1) == "")) end

    assert {:ok, rows} = WorksheetParser.parse(sheet_xml, workbook, skip_row?: ignore_trimmed)
    assert [["-", "-", "-", "-"]] == rows

    ignore_trimmed_or_dashes = fn row -> ignore_trimmed.(row) or Enum.all?(row, &(&1 == "-")) end

    assert {:ok, rows} =
             WorksheetParser.parse(sheet_xml, workbook, skip_row?: ignore_trimmed_or_dashes)

    assert [] == rows
  end

  test "should return cell structs instead of values when cell_data_format is :cell" do
    {:ok, package} = XlsxReader.open(TestFixtures.path("has_formulas.xlsx"))
    {:ok, sheets} = XlsxReader.sheets(package, cell_data_format: :cell)

    expected = [
      {"sheet_1",
       [
         [
           %Cell{value: "abc", formula: nil, ref: "A1"},
           %Cell{value: 123.0, formula: nil, ref: "B1"}
         ]
       ]},
      {"sheet_2",
       [
         [
           %Cell{value: "def", formula: nil, ref: "A1"},
           %Cell{value: 456.0, formula: nil, ref: "B1"},
           "",
           %Cell{value: 466.0, formula: "SUM(B1, 10)", ref: "D1"}
         ]
       ]}
    ]

    assert expected == sheets
  end

  test "should return shared formulas as part of Cell struct", %{workbook: workbook} do
    sheet_xml =
      TestFixtures.read!("xml/worhseetWithSharedFormulas.xml")
      |> String.replace("\n", "")
      |> String.replace("\t", "")

    expected = [
      [
        %XlsxReader.Cell{value: "1", formula: nil, ref: "A1"},
        %XlsxReader.Cell{value: "6", formula: "SUM(A1:A3)", ref: "B1"}
      ],
      [
        %XlsxReader.Cell{value: "2", formula: nil, ref: "A2"},
        %XlsxReader.Cell{value: "6", formula: "SUM(A1:A3)", ref: "B2"}
      ],
      [
        %XlsxReader.Cell{value: "3", formula: nil, ref: "A3"},
        %XlsxReader.Cell{value: "6", formula: "SUM(A1:A3)", ref: "B3"}
      ]
    ]

    assert {:ok, expected} == WorksheetParser.parse(sheet_xml, workbook, cell_data_format: :cell)
  end

  test "should include or exclude hidden sheets based on an option" do
    filepath = TestFixtures.path("hidden_sheets.xlsx")

    {:ok, package} = XlsxReader.open(filepath, exclude_hidden_sheets?: false)
    all_sheet_names = package |> XlsxReader.sheet_names()
    assert all_sheet_names == ["Sheet 1", "Sheet 2", "Sheet 3"]

    {:ok, package} = XlsxReader.open(filepath, exclude_hidden_sheets?: true)
    visible_sheet_names = package |> XlsxReader.sheet_names()
    assert visible_sheet_names == ["Sheet 1"]
  end
end

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

  test "returns custom blank_value for empty <v/> cells when type_conversion is disabled",
       %{workbook: workbook} do
    sheet_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" t="n"><v></v></c>
        </row>
      </sheetData>
    </worksheet>
    """

    assert {:ok, [[:custom_blank]]} =
             WorksheetParser.parse(sheet_xml, workbook,
               type_conversion: false,
               blank_value: :custom_blank
             )
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

  test "should return empty value when result of formulas is not set", %{workbook: workbook} do
    sheet_xml =
      TestFixtures.read!("xml/worksheetWithFormulasWithoutValue.xml")
      |> String.replace("\n", "")
      |> String.replace("\t", "")

    expected = [
      [
        %Cell{value: "1", formula: nil, ref: "A1"},
        %Cell{value: "", formula: "SUM(A1:A3)", ref: "B1"}
      ]
    ]

    assert {:ok, expected} == WorksheetParser.parse(sheet_xml, workbook, cell_data_format: :cell)
  end

  test "should return shared formulas as part of Cell struct", %{workbook: workbook} do
    sheet_xml =
      TestFixtures.read!("xml/worksheetWithSharedFormulas.xml")
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

  test "should return empty value when result of shared formulas is not set", %{
    workbook: workbook
  } do
    sheet_xml =
      TestFixtures.read!("xml/worksheetWithSharedFormulasWithoutValue.xml")
      |> String.replace("\n", "")
      |> String.replace("\t", "")

    expected = [
      [
        %XlsxReader.Cell{value: "1", formula: nil, ref: "A1"},
        %XlsxReader.Cell{value: "", formula: "SUM(A1:A3)", ref: "B1"}
      ],
      [
        %XlsxReader.Cell{value: "2", formula: nil, ref: "A2"},
        %XlsxReader.Cell{value: "", formula: "SUM(A1:A3)", ref: "B2"}
      ],
      [
        %XlsxReader.Cell{value: "3", formula: nil, ref: "A3"},
        %XlsxReader.Cell{value: "", formula: "SUM(A1:A3)", ref: "B3"}
      ]
    ]

    assert {:ok, expected} ==
             WorksheetParser.parse(sheet_xml, workbook,
               cell_data_format: :cell,
               type_conversion: true
             )
  end

  describe "limits" do
    defp sheet_xml(cell_refs_by_row) do
      rows =
        for {row, refs} <- cell_refs_by_row do
          cells = for ref <- refs, do: ~s(<c r="#{ref}" t="n"><v>1</v></c>)
          ~s(<row r="#{row}">#{cells}</row>)
        end

      ~s(<worksheet><sheetData>#{rows}</sheetData></worksheet>)
    end

    defp full_sheet do
      sheet_xml([{1, ~w(A1 B1 C1)}, {2, ~w(A2 B2 C2)}, {3, ~w(A3 B3 C3)}])
    end

    test "accepts sheets within the limits", %{workbook: workbook} do
      assert {:ok, [[_, _, _], [_, _, _], [_, _, _]]} =
               WorksheetParser.parse(full_sheet(), workbook,
                 max_rows: 3,
                 max_columns: 3,
                 max_cells: 9
               )
    end

    test "rejects sheets exceeding max_rows", %{workbook: workbook} do
      assert WorksheetParser.parse(full_sheet(), workbook, max_rows: 2) ==
               {:error, "row limit exceeded"}
    end

    test "rejects sheets exceeding max_columns", %{workbook: workbook} do
      assert WorksheetParser.parse(full_sheet(), workbook, max_columns: 2) ==
               {:error, "column limit exceeded"}
    end

    test "rejects sheets exceeding max_cells", %{workbook: workbook} do
      assert WorksheetParser.parse(full_sheet(), workbook, max_cells: 8) ==
               {:error, "cell limit exceeded"}
    end

    test "the lowest limit wins", %{workbook: workbook} do
      assert WorksheetParser.parse(full_sheet(), workbook,
               max_rows: 100,
               max_columns: 100,
               max_cells: 8
             ) == {:error, "cell limit exceeded"}

      assert WorksheetParser.parse(full_sheet(), workbook,
               max_rows: 2,
               max_columns: 100,
               max_cells: 100
             ) == {:error, "row limit exceeded"}

      assert WorksheetParser.parse(full_sheet(), workbook,
               max_rows: 100,
               max_columns: 2,
               max_cells: 100
             ) == {:error, "column limit exceeded"}
    end

    test "rejects sparse sheets without expanding them", %{workbook: workbook} do
      sparse = sheet_xml([{1_048_576, ["XFD1048576"]}])

      assert WorksheetParser.parse(sparse, workbook, max_rows: 1000) ==
               {:error, "row limit exceeded"}

      assert WorksheetParser.parse(sparse, workbook, max_columns: 1000) ==
               {:error, "column limit exceeded"}

      assert WorksheetParser.parse(sparse, workbook, max_cells: 1_000_000) ==
               {:error, "cell limit exceeded"}
    end

    test "counts blank cells created for omitted cells", %{workbook: workbook} do
      xml = sheet_xml([{1, ["C1"]}])

      assert {:ok, [["", "", _]]} = WorksheetParser.parse(xml, workbook, max_cells: 3)

      assert WorksheetParser.parse(xml, workbook, max_cells: 2) == {:error, "cell limit exceeded"}
    end

    test "counts blank rows created for omitted rows", %{workbook: workbook} do
      xml = sheet_xml([{1, ["A1"]}, {3, ["A3"]}])

      assert {:ok, [[_], [""], [_]]} = WorksheetParser.parse(xml, workbook, max_cells: 3)

      assert WorksheetParser.parse(xml, workbook, max_cells: 2) == {:error, "cell limit exceeded"}
    end

    test "does not count omitted rows that are not included", %{workbook: workbook} do
      xml = sheet_xml([{1, ["A1"]}, {3, ["A3"]}])

      assert {:ok, [[_], [_]]} =
               WorksheetParser.parse(xml, workbook, empty_rows: false, max_cells: 2)

      assert WorksheetParser.parse(xml, workbook, empty_rows: false, max_rows: 2) ==
               {:error, "row limit exceeded"}
    end
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

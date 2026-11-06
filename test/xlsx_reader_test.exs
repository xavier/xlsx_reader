defmodule XlsxReaderTest do
  use ExUnit.Case

  test "README install version check" do
    app = :xlsx_reader

    app_version = "#{Application.spec(app, :vsn)}"
    readme = File.read!("README.md")
    [_, readme_versions] = Regex.run(~r/{:#{app}, "(.+)"}/, readme)

    assert Version.match?(
             app_version,
             readme_versions
           ),
           """
           Install version constraint in README.md does not match to current app version.
           Current App Version: #{app_version}
           Readme Install Versions: #{readme_versions}
           """
  end

  describe "open/2" do
    test "opens a xlsx file from the file system" do
      xlsx = TestFixtures.path("test.xlsx")

      assert {:ok, %XlsxReader.Package{}} = XlsxReader.open(xlsx)
      assert {:ok, %XlsxReader.Package{}} = XlsxReader.open(xlsx, source: :path)
    end

    test "open a xlsx file from memory" do
      xlsx = TestFixtures.read!("test.xlsx")

      assert {:ok, %XlsxReader.Package{}} = XlsxReader.open(xlsx, source: :binary)
    end

    test "rejects non-xlsx file" do
      xlsx = TestFixtures.path("test.zip")

      assert {:error, "invalid xlsx file"} = XlsxReader.open(xlsx)
    end

    test "rejects non-zip file" do
      xlsx = TestFixtures.path("not_a_zip.zip")

      assert {:error, "invalid zip file"} = XlsxReader.open(xlsx)
    end

    test "rejects relative and absolute path to directory" do
      relative_path = "test"
      absolute_path = Path.absname(relative_path)

      assert {:error, "invalid zip file"} = XlsxReader.open(relative_path)
      assert {:error, "invalid zip file"} = XlsxReader.open(absolute_path)
    end

    test "supported custom formats" do
      xlsx = TestFixtures.path("test.xlsx")

      assert {:ok, package} =
               XlsxReader.open(xlsx,
                 supported_custom_formats: [
                   {"[$CHF]0.00", :string}
                 ]
               )

      {:ok, sheet} = XlsxReader.sheet(package, "Sheet 3")

      assert [
               ["", _],
               ["date", _],
               ["datetime", _],
               ["time", _],
               ["percentage", _],
               ["money chf", "100"],
               ["money usd", _],
               ["ticked", _],
               ["not ticked", _],
               ["hyperlink", _]
             ] = sheet
    end
  end

  describe "sheet_names/1" do
    setup do
      {:ok, package} = XlsxReader.open(TestFixtures.path("test.xlsx"))

      {:ok, %{package: package}}
    end

    test "lists the sheets in workbook", %{package: package} do
      assert ["Sheet 1", "Sheet 2", "Sheet 3"] == XlsxReader.sheet_names(package)
    end
  end

  describe "sheet/3" do
    setup do
      {:ok, package} = XlsxReader.open(TestFixtures.path("test.xlsx"))

      {:ok, %{package: package}}
    end

    test "returns the contents of the sheet by name", %{package: package} do
      assert {:ok,
              [
                ["A", "B", "C" | _],
                [1.0, 2.0, 3.0 | _]
                | _
              ]} = XlsxReader.sheet(package, "Sheet 1")
    end

    test "type conversion off", %{package: package} do
      assert {:ok,
              [
                _,
                ["date", "43784"]
                | _
              ]} = XlsxReader.sheet(package, "Sheet 3", type_conversion: false)
    end

    test "number type", %{package: package} do
      assert {:ok,
              [
                ["A", "B", "C" | _],
                [
                  %Decimal{coef: 1, exp: 0, sign: 1},
                  %Decimal{coef: 2, exp: 0, sign: 1},
                  %Decimal{coef: 3, exp: 0, sign: 1} | _
                ]
                | _
              ]} = XlsxReader.sheet(package, "Sheet 1", number_type: Decimal)
    end

    test "custom blank value", %{package: package} do
      assert {:ok,
              [
                ["n/a", "n/a"]
                | _
              ]} = XlsxReader.sheet(package, "Sheet 3", blank_value: "n/a")
    end

    test "skip empty rows", %{package: package} do
      assert {:ok,
              [
                ["date" | _]
                | _
              ]} = XlsxReader.sheet(package, "Sheet 3", empty_rows: false)
    end
  end

  describe "sheets/2" do
    setup do
      {:ok, package} = XlsxReader.open(TestFixtures.path("test.xlsx"))

      {:ok, %{package: package}}
    end

    test "load all sheets", %{package: package} do
      assert {:ok,
              [
                {"Sheet 1", [["A", "B", "C" | _] | _]},
                {"Sheet 2", [["", "", "", "", ""] | _]},
                {"Sheet 3", [["", ""] | _]}
              ]} = XlsxReader.sheets(package)
    end

    test "filters sheets", %{package: package} do
      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.sheets(package, only: ["Sheet 1", "Sheet 3"])

      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.sheets(package, only: [~r/Sheet [13]/])

      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.sheets(package, except: "Sheet 2")

      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.sheets(package, only: ~r/Sheet \d+/, except: ["Sheet 2"])
    end
  end

  describe "async_sheets/3" do
    setup do
      {:ok, package} = XlsxReader.open(TestFixtures.path("test.xlsx"))

      {:ok, %{package: package}}
    end

    test "load all sheets", %{package: package} do
      assert {:ok,
              [
                {"Sheet 1", [["A", "B", "C" | _] | _]},
                {"Sheet 2", [["", "", "", "", ""] | _]},
                {"Sheet 3", [["", ""] | _]}
              ]} = XlsxReader.async_sheets(package)
    end

    test "filters sheets", %{package: package} do
      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.async_sheets(package, only: ["Sheet 1", "Sheet 3"])

      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.async_sheets(package, only: [~r/Sheet [13]/])

      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.async_sheets(package, except: "Sheet 2")

      assert {:ok,
              [
                {"Sheet 1", _},
                {"Sheet 3", _}
              ]} = XlsxReader.async_sheets(package, only: ~r/Sheet \d+/, except: ["Sheet 2"])
    end

    test "returns the underlying error tuple when a sheet fails to load", %{package: package} do
      bogus_sheet = %XlsxReader.Sheet{name: "Bogus", rid: "rIdMissing"}
      bogus_package = %{package | workbook: %{package.workbook | sheets: [bogus_sheet]}}

      assert {:error, "sheet relationship not found"} = XlsxReader.async_sheets(bogus_package)
    end
  end

  describe "max_uncompressed_size" do
    @content_types ~s(<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>)
    @workbook ~s(<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Bomb" r:id="rId1" sheetId="1"/></sheets></workbook>)
    @relationship_type "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    defp xlsx(sheet, shared_strings \\ nil) do
      shared_strings_rel =
        if shared_strings,
          do:
            ~s(<Relationship Id="rId2" Target="sharedStrings.xml" Type="#{@relationship_type}/sharedStrings"/>),
          else: ""

      files =
        [
          {~c"[Content_Types].xml", @content_types},
          {~c"xl/workbook.xml", @workbook},
          {~c"xl/_rels/workbook.xml.rels",
           ~s(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="#{@relationship_type}/worksheet"/>#{shared_strings_rel}</Relationships>)},
          {~c"xl/worksheets/sheet1.xml", sheet}
        ] ++ if(shared_strings, do: [{~c"xl/sharedStrings.xml", shared_strings}], else: [])

      {:ok, {_, binary}} = :zip.create(~c"bomb.xlsx", files, [:memory, {:compress, :all}])
      binary
    end

    # ~25MB of XML, compressed to a few hundred KB
    defp bomb_sheet do
      cells = String.duplicate(~s(<c t="n"><v>1</v></c>), 1_000_000)
      ~s(<worksheet><sheetData><row r="1">#{cells}</row></sheetData></worksheet>)
    end

    test "rejects sheets inflating over the limit" do
      binary = xlsx(bomb_sheet())

      assert byte_size(binary) < 500_000

      {:ok, package} = XlsxReader.open(binary, source: :binary, max_uncompressed_size: 1_000_000)

      assert XlsxReader.sheet(package, "Bomb") == {:error, "uncompressed size limit exceeded"}
      assert XlsxReader.sheets(package) == {:error, "uncompressed size limit exceeded"}
      assert XlsxReader.async_sheets(package) == {:error, "uncompressed size limit exceeded"}
    end

    test "rejects shared strings inflating over the limit when opening" do
      shared_strings =
        ~s(<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">#{String.duplicate("<si><t>a</t></si>", 1_000_000)}</sst>)

      binary = xlsx(~s(<worksheet><sheetData/></worksheet>), shared_strings)

      assert XlsxReader.open(binary, source: :binary, max_uncompressed_size: 1_000_000) ==
               {:error, "uncompressed size limit exceeded"}
    end

    test "does not limit sheets by default and within the limit" do
      sheet =
        ~s(<worksheet><sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData></worksheet>)

      binary = xlsx(sheet)

      for options <- [[], [max_uncompressed_size: 1_000_000]] do
        {:ok, package} = XlsxReader.open(binary, [source: :binary] ++ options)

        assert XlsxReader.sheet(package, "Bomb") == {:ok, [[1.0]]}
      end
    end
  end
end

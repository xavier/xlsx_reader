defmodule XlsxReaderTest do
  use ExUnit.Case

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
  end
end

# credo:disable-for-this-file Credo.Check.Readability.LargeNumbers
defmodule CompatibilityTest do
  use ExUnit.Case

  test "google_spreadsheet.xlsx" do
    assert {:ok, package} = XlsxReader.open(TestFixtures.path("google_spreadsheet.xlsx"))

    assert ["Sheet1"] = XlsxReader.sheet_names(package)

    assert {:ok,
            [
              ["integer", 123.0],
              ["float", 123.456],
              ["percentage", 12.5],
              ["date", 43784.0],
              ["time", ~N[1899-12-30 11:45:00]],
              ["ticked\n", true],
              ["unticked", false],
              ["image", ""]
            ]} = XlsxReader.sheet(package, "Sheet1")
  end

  test "merged.xlsx" do
    assert {:ok, package} = XlsxReader.open(TestFixtures.path("merged.xlsx"))

    assert ["merged"] = XlsxReader.sheet_names(package)

    assert {:ok,
            [
              ["horizontal", "", "vertical"],
              ["horizontal + vertical", "", ""],
              ["", "", "none"]
            ]} = XlsxReader.sheet(package, "merged")
  end
end

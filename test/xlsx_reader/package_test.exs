defmodule XlsxReader.PackageTest do
  use ExUnit.Case

  alias XlsxReader.Package

  describe "open/1" do
    test "opens a xlsx file" do
      zip_handle = {:path, TestFixtures.path("test.xlsx")}

      assert {:ok, package} = Package.open(zip_handle)

      assert package.zip_handle == zip_handle
    end

    test "rejects non-xlsx file" do
      assert {:error, "invalid xlsx file"} = Package.open({:path, TestFixtures.path("test.zip")})
    end

    test "rejects non-zip file" do
      assert {:error, "invalid zip file"} =
               Package.open({:path, TestFixtures.path("not_a_zip.zip")})
    end
  end

  describe "load_sheet/2" do
    setup do
      {:ok, package} = Package.open({:path, TestFixtures.path("test.xlsx")})

      {:ok, %{package: package}}
    end

    test "loads a sheet by name", %{package: package} do
      assert {:ok,
              [
                ["Table 1", "", "" | _],
                ["A", "B", "C" | _],
                ["1", "2", "3" | _],
                ["2", "4", "6" | _]
                | _
              ]} = Package.load_sheet(package, "Sheet 1")

      assert {:ok,
              [
                ["Table 1", "" | _],
                ["", "" | _],
                ["some ", "test" | _]
                | _
              ]} = Package.load_sheet(package, "Sheet 2")
    end
  end

  describe "load_sheets/1" do
    setup do
      {:ok, package} = Package.open({:path, TestFixtures.path("test.xlsx")})

      {:ok, %{package: package}}
    end

    test "load all sheets", %{package: package} do
      assert [
               {"Sheet 1", [["Table 1", "", "" | _], ["A", "B", "C" | _] | _]},
               {"Sheet 2", [["Table 1", "" | _], ["", "" | _] | _]}
             ] = Package.load_sheets(package)
    end
  end
end

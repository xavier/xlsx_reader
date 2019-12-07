defmodule XlsxReader.PackageLoaderTest do
  use ExUnit.Case

  alias XlsxReader.{PackageLoader, Unzip}

  describe "open/1" do
    test "opens a xlsx file" do
      zip_handle = Unzip.handle(TestFixtures.path("test.xlsx"), :path)

      assert {:ok, package} = PackageLoader.open(zip_handle)

      assert package.zip_handle == zip_handle
    end

    test "rejects non-xlsx file" do
      zip_handle = Unzip.handle(TestFixtures.path("test.zip"), :path)

      assert {:error, "invalid xlsx file"} = PackageLoader.open(zip_handle)
    end

    test "rejects non-zip file" do
      zip_handle = Unzip.handle(TestFixtures.path("not_a_zip.zip"), :path)

      assert {:error, "invalid zip file"} = PackageLoader.open(zip_handle)
    end
  end

  describe "load_sheet_by_name/2" do
    setup do
      zip_handle = Unzip.handle(TestFixtures.path("test.xlsx"), :path)
      {:ok, package} = PackageLoader.open(zip_handle)

      {:ok, %{package: package}}
    end

    test "loads a sheet by name", %{package: package} do
      assert {:ok,
              [
                ["A", "B", "C" | _],
                [1.0, 2.0, 3.0 | _],
                [2.0, 4.0, 6.0 | _]
                | _
              ]} = PackageLoader.load_sheet_by_name(package, "Sheet 1")

      assert {:ok,
              [
                ["", "" | _],
                ["some ", "test" | _]
                | _
              ]} = PackageLoader.load_sheet_by_name(package, "Sheet 2")
    end
  end
end

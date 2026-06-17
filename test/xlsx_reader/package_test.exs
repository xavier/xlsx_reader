defmodule XlsxReader.PackageLoaderTest do
  use ExUnit.Case

  alias XlsxReader.{PackageLoader, ZipArchive}

  describe "open/1" do
    test "opens a xlsx file" do
      zip_handle = ZipArchive.handle(TestFixtures.path("test.xlsx"), :path)

      assert {:ok, package} = PackageLoader.open(zip_handle)

      assert package.zip_handle == zip_handle
    end

    test "rejects non-xlsx file" do
      zip_handle = ZipArchive.handle(TestFixtures.path("test.zip"), :path)

      assert {:error, "invalid xlsx file"} = PackageLoader.open(zip_handle)
    end

    test "rejects non-zip file" do
      zip_handle = ZipArchive.handle(TestFixtures.path("not_a_zip.zip"), :path)

      assert {:error, "invalid zip file"} = PackageLoader.open(zip_handle)
    end

    # https://github.com/xavier/xlsx_reader/issues/46
    test "resolves package-root-absolute relationship targets" do
      # The `package` fixture uses targets such as "/xl/styles.xml" and
      # "/xl/worksheets/sheet3.xml", i.e. absolute paths that already include
      # the "xl/" segment. These must not be double-prefixed into "xl/xl/...".
      zip_handle = ZipArchive.handle(build_package_zip(), :binary)

      assert {:ok, %XlsxReader.Package{} = package} = PackageLoader.open(zip_handle)

      # "Sheet 3" relates to "/xl/worksheets/sheet3.xml"
      assert {:ok, [_ | _]} = PackageLoader.load_sheet_by_name(package, "Sheet 3")
    end

    # https://github.com/xavier/xlsx_reader/issues/50
    test "returns an error when the styles part is referenced but missing" do
      zip_handle = ZipArchive.handle(build_package_zip(exclude: ["xl/styles.xml"]), :binary)

      assert {:error, "file \"xl/styles.xml\" not found in archive"} =
               PackageLoader.open(zip_handle)
    end

    # https://github.com/xavier/xlsx_reader/issues/50
    test "returns an error when the shared strings part is referenced but missing" do
      zip_handle = ZipArchive.handle(build_package_zip(exclude: ["xl/sharedStrings.xml"]), :binary)

      assert {:error, "file \"xl/sharedStrings.xml\" not found in archive"} =
               PackageLoader.open(zip_handle)
    end
  end

  describe "load_sheet_by_name/2" do
    setup do
      zip_handle = ZipArchive.handle(TestFixtures.path("test.xlsx"), :path)
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

  # Builds an in-memory .xlsx archive from the unpacked `package` fixture so the
  # relationship targets defined in its workbook.xml.rels are exercised end-to-end.
  defp build_package_zip(opts \\ []) do
    root = TestFixtures.path("package")
    exclude = Keyword.get(opts, :exclude, [])

    files =
      root
      |> Path.join("**")
      |> Path.wildcard()
      |> Enum.reject(&File.dir?/1)
      |> Enum.map(fn path -> {Path.relative_to(path, root), File.read!(path)} end)
      |> Enum.reject(fn {relative_path, _} -> relative_path in exclude end)
      |> Enum.map(fn {relative_path, contents} -> {String.to_charlist(relative_path), contents} end)

    {:ok, {_name, zip}} = :zip.create(~c"package.xlsx", files, [:memory])
    zip
  end
end

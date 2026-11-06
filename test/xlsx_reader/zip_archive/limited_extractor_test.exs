defmodule XlsxReader.ZipArchive.LimitedExtractorTest do
  use ExUnit.Case

  alias XlsxReader.ZipArchive

  # `ZipArchive.extract/3` delegates to the limited extractor when given a limit

  describe "extract/3 with a size limit" do
    @limit_error {:error, "uncompressed size limit exceeded"}

    @central_directory_magic <<0x02014B50::32-little>>

    defp zip(contents, options \\ []) do
      {:ok, {_, binary}} =
        :zip.create(~c"test.zip", [{~c"file.xml", contents}], [:memory | options])

      binary
    end

    # Overwrites the uncompressed size declared in the central directory
    defp forge_declared_size(zip, size) do
      {position, _} = :binary.match(zip, @central_directory_magic)
      <<before::binary-size(position + 24), _::32, rest::binary>> = zip
      <<before::binary, size::32-little, rest::binary>>
    end

    test "extracts files within the limit" do
      contents = String.duplicate("<c/>", 1000)

      for options <- [[], [{:compress, :all}], [{:uncompress, :all}]] do
        zip_handle = ZipArchive.handle(zip(contents, options), :binary)

        assert {:ok, ^contents} = ZipArchive.extract(zip_handle, "file.xml", 4000)
        assert {:ok, ^contents} = ZipArchive.extract(zip_handle, "file.xml", 1_000_000)
      end
    end

    test "rejects files exceeding the limit" do
      contents = String.duplicate("<c/>", 1000)

      for options <- [[{:compress, :all}], [{:uncompress, :all}]] do
        zip_handle = ZipArchive.handle(zip(contents, options), :binary)

        assert ZipArchive.extract(zip_handle, "file.xml", 3999) == @limit_error
      end
    end

    test "rejects deflated files exceeding the limit even if the declared size is forged" do
      contents = String.duplicate("<c/>", 1_000_000)
      forged = contents |> zip([{:compress, :all}]) |> forge_declared_size(10)
      zip_handle = ZipArchive.handle(forged, :binary)

      assert ZipArchive.extract(zip_handle, "file.xml", 1000) == @limit_error
    end

    test "rejects stored files exceeding the limit even if the declared size is forged" do
      contents = String.duplicate("<c/>", 1000)
      forged = contents |> zip([{:uncompress, :all}]) |> forge_declared_size(10)
      zip_handle = ZipArchive.handle(forged, :binary)

      assert ZipArchive.extract(zip_handle, "file.xml", 1000) == @limit_error
    end

    test "rejects files with a bad checksum" do
      zip = zip("<c/>", [{:compress, :all}])
      # CRC-32 field of the local file header
      <<before::binary-size(14), crc::32-little, rest::binary>> = zip
      corrupt = <<before::binary, Bitwise.bxor(crc, 1)::32-little, rest::binary>>

      assert ZipArchive.extract(ZipArchive.handle(corrupt, :binary), "file.xml", 1000) ==
               {:error, "invalid zip file"}
    end

    test "rejects truncated archives" do
      zip = zip(String.duplicate("<c/>", 1000), [{:compress, :all}])
      truncated = binary_part(zip, 0, 40)

      assert {:error, "invalid zip file"} =
               ZipArchive.extract(ZipArchive.handle(truncated, :binary), "file.xml", 100_000)
    end

    test "returns the same contents as extract/2 for the files of a real xlsx file" do
      for source <- [
            {:path, TestFixtures.path("test.xlsx")},
            {:binary, TestFixtures.read!("test.xlsx")}
          ] do
        zip_handle = ZipArchive.handle(elem(source, 1), elem(source, 0))
        {:ok, files} = ZipArchive.list(zip_handle)

        assert files != []

        for file <- files do
          assert ZipArchive.extract(zip_handle, file, 10_000_000) ==
                   ZipArchive.extract(zip_handle, file)
        end
      end
    end

    @tag :tmp_dir
    test "extracts from a file on disk", %{tmp_dir: tmp_dir} do
      path = Path.join(tmp_dir, "test.zip")
      contents = String.duplicate("<c/>", 1000)
      File.write!(path, zip(contents, [{:compress, :all}]))
      zip_handle = ZipArchive.handle(path, :path)

      assert {:ok, ^contents} = ZipArchive.extract(zip_handle, "file.xml", 4000)
      assert ZipArchive.extract(zip_handle, "file.xml", 100) == @limit_error

      assert ZipArchive.extract(zip_handle, "bogus.xml", 100) ==
               {:error, "file \"bogus.xml\" not found in archive"}
    end

    test "invalid zip file and missing zip file" do
      assert ZipArchive.extract(
               ZipArchive.handle(TestFixtures.path("not_a_zip.zip"), :path),
               "a",
               100
             ) ==
               {:error, "invalid zip file"}

      assert ZipArchive.extract(ZipArchive.handle("__does_not_exist__", :path), "a", 100) ==
               {:error, "file not found"}
    end
  end
end

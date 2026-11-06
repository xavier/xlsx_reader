defmodule XlsxReader.ZipArchive.SizeLimit do
  @moduledoc false

  # Protection against "zip bombs".
  #
  # `:zip.extract/2` inflates a whole file in memory with no limit, and the sizes
  # declared in the archive can be forged. Before extracting, this inflates the
  # file once, counting and discarding the output, and stops at the limit.
  #
  # Stored and deflated files only (as `:zip`); zip64 is untested.

  @local_file_header_size 30
  @local_file_magic 0x04034B50
  @stored 0
  @deflated 8

  @spec check(XlsxReader.ZipArchive.zip_handle(), String.t(), non_neg_integer()) ::
          :ok | {:error, term()}
  def check(zip_handle, file, max_size) do
    with {:ok, entries} <- :zip.list_dir(zip_source(zip_handle)),
         {:ok, {:zip_file, _, _, _, offset, comp_size}} <- find_entry(entries, file),
         {:ok, header} <- read(zip_handle, offset, @local_file_header_size),
         {:ok, method, data_offset} <- parse_local_header(header, offset) do
      check_size(method, zip_handle, data_offset, comp_size, max_size)
    end
  end

  defp zip_source({:path, path}), do: String.to_charlist(path)
  defp zip_source({:binary, binary}), do: binary

  defp find_entry(entries, file) do
    case Enum.find(
           entries,
           &(match?({:zip_file, _, _, _, _, _}, &1) and to_string(elem(&1, 1)) == file)
         ) do
      nil -> {:error, "file #{inspect(file)} not found in archive"}
      entry -> {:ok, entry}
    end
  end

  # Encrypted files (flag bit 0) are not supported
  defp parse_local_header(
         <<@local_file_magic::32-little, _version::16, flags::16-little, method::16-little,
           _time_date_crc_and_sizes::128, name_length::16-little, extra_length::16-little>>,
         offset
       )
       when Bitwise.band(flags, 1) == 0 do
    {:ok, method, offset + @local_file_header_size + name_length + extra_length}
  end

  defp parse_local_header(_header, _offset), do: invalid()

  # Stored data isn't compressed: its size is the size of the data in the archive
  defp check_size(@stored, _zip_handle, _offset, size, max_size),
    do: within_limit(size, max_size)

  defp check_size(@deflated, zip_handle, offset, size, max_size) do
    with {:ok, compressed} <- read(zip_handle, offset, size) do
      z = :zlib.open()

      try do
        :ok = :zlib.inflateInit(z, -15)
        count_inflated(z, compressed, max_size, 0)
      rescue
        ErlangError -> invalid()
      after
        :zlib.close(z)
      end
    end
  end

  defp check_size(_method, _zip_handle, _offset, _size, _max_size), do: invalid()

  defp count_inflated(z, input, max_size, size) do
    {status, output} = :zlib.safeInflate(z, input)
    size = size + IO.iodata_length(output)

    case within_limit(size, max_size) do
      :ok when status == :continue -> count_inflated(z, [], max_size, size)
      result -> result
    end
  end

  defp within_limit(size, max_size) when size > max_size,
    do: {:error, "uncompressed size limit exceeded"}

  defp within_limit(_size, _max_size), do: :ok

  defp read({:binary, binary}, offset, size) when offset + size <= byte_size(binary),
    do: {:ok, binary_part(binary, offset, size)}

  defp read({:binary, _binary}, _offset, _size), do: invalid()

  defp read({:path, path}, offset, size) do
    with {:ok, io_device} <- :file.open(path, [:read, :raw, :binary]) do
      try do
        case :file.pread(io_device, offset, size) do
          {:ok, data} when byte_size(data) == size -> {:ok, data}
          _ -> invalid()
        end
      after
        :file.close(io_device)
      end
    end
  end

  defp invalid, do: {:error, "invalid zip file"}
end

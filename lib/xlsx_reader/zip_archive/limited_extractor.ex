defmodule XlsxReader.ZipArchive.LimitedExtractor do
  @moduledoc false

  # EXPERIMENTAL: extraction of a file from a zip archive with a limit on its
  # uncompressed size, as a protection against "zip bombs".
  #
  # `:zip.extract/2` inflates a whole file in memory without any limit and
  # doesn't verify the sizes declared in the archive, which an attacker can forge.
  # Here the file is located using the central directory (`:zip.list_dir/1`),
  # then inflated in chunks, giving up as soon as the limit is exceeded.
  #
  # Only the "stored" and "deflated" compression methods are supported (as in
  # `:zip`). Encrypted files are rejected. Zip64 archives are untested.
  #
  # Returns `{:ok, binary}`, `{:error, message}` or the raw error tuple of
  # `:zip`/`:file` for the caller to translate.

  require Record
  Record.defrecordp(:file_info, Record.extract(:file_info, from_lib: "kernel/include/file.hrl"))

  @local_file_magic 0x04034B50
  @data_descriptor_magic 0x08074B50
  @stored 0
  @deflated 8
  @read_chunk_size 65_536

  @spec extract(XlsxReader.ZipArchive.zip_handle(), String.t(), non_neg_integer()) ::
          {:ok, binary()} | {:error, term()}
  def extract(zip_handle, file, max_size) do
    with {:ok, entries} <- :zip.list_dir(zip_source(zip_handle)),
         {:ok, entry} <- find_entry(entries, file),
         :ok <- check_declared_size(entry, max_size) do
      read_entry(zip_handle, entry, max_size)
    end
  end

  defp zip_source({:path, path}), do: String.to_charlist(path)
  defp zip_source({:binary, binary}), do: binary

  defp find_entry(entries, file) do
    entries
    |> Enum.find(fn
      {:zip_file, name, _info, _comment, _offset, _comp_size} -> to_string(name) == file
      _ -> false
    end)
    |> case do
      nil -> {:error, "file #{inspect(file)} not found in archive"}
      entry -> {:ok, entry}
    end
  end

  defp check_declared_size({:zip_file, _, info, _, _, _}, max_size) do
    if file_info(info, :size) > max_size, do: {:error, size_limit_error()}, else: :ok
  end

  defp read_entry(zip_handle, {:zip_file, _, _, _, offset, comp_size}, max_size) do
    with_device(zip_handle, &read_entry_data(&1, offset, comp_size, max_size))
  end

  defp read_entry_data(device, offset, comp_size, max_size) do
    with {:ok, header} <- read_exact(device, offset, 30),
         {:ok, entry} <- parse_local_header(header),
         data_pos = offset + 30 + entry.name_length + entry.extra_length,
         {:ok, data, crc} <- read_data(device, entry.method, data_pos, comp_size, max_size),
         {:ok, expected_crc} <- expected_crc(device, entry, data_pos + comp_size),
         :ok <- verify_crc(crc, expected_crc) do
      {:ok, data}
    end
  end

  defp parse_local_header(
         <<@local_file_magic::32-little, _version::16, flags::16-little, method::16-little,
           _time_and_date::32, crc::32-little, _sizes::64, name_length::16-little,
           extra_length::16-little>>
       ) do
    if Bitwise.band(flags, 1) == 1 do
      # Encrypted files are not supported
      {:error, "invalid zip file"}
    else
      {:ok,
       %{
         method: method,
         crc: crc,
         data_descriptor?: Bitwise.band(flags, 8) == 8,
         name_length: name_length,
         extra_length: extra_length
       }}
    end
  end

  defp parse_local_header(_), do: {:error, "invalid zip file"}

  # CRC is stored after the data when the sizes weren't known when the header was written
  defp expected_crc(_device, %{data_descriptor?: false, crc: crc}, _pos), do: {:ok, crc}

  defp expected_crc(device, %{data_descriptor?: true}, pos) do
    case read_exact(device, pos, 8) do
      {:ok, <<@data_descriptor_magic::32-little, crc::32-little>>} -> {:ok, crc}
      {:ok, <<crc::32-little, _::32>>} -> {:ok, crc}
      error -> error
    end
  end

  defp verify_crc(crc, crc), do: :ok
  defp verify_crc(_, _), do: {:error, "invalid zip file"}

  defp read_data(device, @stored, pos, size, max_size) do
    if size > max_size do
      {:error, size_limit_error()}
    else
      with {:ok, data} <- read_exact(device, pos, size) do
        {:ok, data, :erlang.crc32(data)}
      end
    end
  end

  defp read_data(device, @deflated, pos, size, max_size) do
    z = :zlib.open()

    try do
      :ok = :zlib.inflateInit(z, -15)
      inflate(z, device, pos, size, max_size, %{size: 0, crc: 0, chunks: []})
    rescue
      ErlangError -> {:error, "invalid zip file"}
    after
      :zlib.close(z)
    end
  end

  defp read_data(_device, _method, _pos, _size, _max_size), do: {:error, "invalid zip file"}

  defp inflate(_z, _device, _pos, 0, _max_size, acc) do
    {:ok, acc.chunks |> Enum.reverse() |> IO.iodata_to_binary(), acc.crc}
  end

  defp inflate(z, device, pos, remaining, max_size, acc) do
    chunk_size = min(@read_chunk_size, remaining)

    with {:ok, chunk} <- read_exact(device, pos, chunk_size),
         {:ok, acc} <- drain(z, chunk, max_size, acc) do
      inflate(z, device, pos + chunk_size, remaining - chunk_size, max_size, acc)
    end
  end

  # `safeInflate/2` returns a limited amount of output at a time:
  # call it until all the input has been consumed.
  defp drain(z, input, max_size, acc) do
    case :zlib.safeInflate(z, input) do
      {:continue, output} ->
        with {:ok, acc} <- accumulate(output, max_size, acc), do: drain(z, [], max_size, acc)

      {:finished, output} ->
        accumulate(output, max_size, acc)
    end
  end

  defp accumulate(output, max_size, acc) do
    size = acc.size + IO.iodata_length(output)

    if size > max_size do
      {:error, size_limit_error()}
    else
      {:ok,
       %{acc | size: size, crc: :erlang.crc32(acc.crc, output), chunks: [output | acc.chunks]}}
    end
  end

  defp size_limit_error, do: "uncompressed size limit exceeded"

  defp with_device({:binary, binary}, fun), do: fun.({:binary, binary})

  defp with_device({:path, path}, fun) do
    case :file.open(path, [:read, :raw, :binary]) do
      {:ok, io_device} ->
        try do
          fun.({:file, io_device})
        after
          :file.close(io_device)
        end

      error ->
        error
    end
  end

  defp read_exact(device, pos, size) do
    case pread(device, pos, size) do
      {:ok, data} when byte_size(data) == size -> {:ok, data}
      _ -> {:error, "invalid zip file"}
    end
  end

  defp pread({:binary, binary}, pos, size) when pos < byte_size(binary),
    do: {:ok, binary_part(binary, pos, min(size, byte_size(binary) - pos))}

  defp pread({:binary, _binary}, _pos, _size), do: :eof
  defp pread({:file, io_device}, pos, size), do: :file.pread(io_device, pos, size)
end

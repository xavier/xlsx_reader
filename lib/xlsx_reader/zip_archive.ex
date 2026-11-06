defmodule XlsxReader.ZipArchive do
  @moduledoc false

  # Zip archive utility functions.
  #
  # To allow for transparent handling of archives located on disk or in memory,
  # you must first obtain a handle with `handle/2` which can then be used to
  # access the contents of the archive.

  alias XlsxReader.ZipArchive.SizeLimit

  @type source :: String.t() | binary()
  @type source_type :: :path | :binary
  @type zip_handle :: {:path, String.t()} | {:binary, binary()}

  @typedoc """
  Maximum size, in bytes, of an extracted file (`:infinity` for no limit)
  """
  @type max_size :: non_neg_integer() | :infinity

  @doc """

  Returns a `zip_handle` to be used by `list/1` and `extract/2`

  """
  @spec handle(source(), source_type()) :: zip_handle()
  def handle(source, type) when type in [:path, :binary],
    do: {type, source}

  @doc """

  Lists the content of the archive.

  """
  @spec list(zip_handle()) :: {:ok, [String.t()]} | XlsxReader.error()
  def list(zip_handle) do
    with {:ok, zip} <- source(zip_handle),
         {:ok, entries} <- :zip.list_dir(zip) do
      {:ok, collect_files(entries)}
    else
      error ->
        translate_zip_error(error)
    end
  end

  @doc """

  Extracts a file from the archive.

  When `max_size` is given, extracting a file whose uncompressed size exceeds it
  returns `{:error, "uncompressed size limit exceeded"}` (see `XlsxReader.ZipArchive.SizeLimit`).

  """
  @spec extract(zip_handle(), String.t(), max_size()) :: {:ok, binary()} | XlsxReader.error()
  def extract(zip_handle, file, max_size \\ :infinity)

  def extract(zip_handle, file, :infinity) do
    with {:ok, zip} <- source(zip_handle),
         {:ok, [{_, contents}]} <- :zip.extract(zip, extract_options(file)) do
      {:ok, contents}
    else
      {:ok, []} ->
        {:error, "file #{inspect(file)} not found in archive"}

      error ->
        translate_zip_error(error)
    end
  end

  def extract(zip_handle, file, max_size) when is_integer(max_size) and max_size >= 0 do
    case SizeLimit.check(zip_handle, file, max_size) do
      :ok -> extract(zip_handle, file, :infinity)
      {:error, message} = error when is_binary(message) -> error
      error -> translate_zip_error(error)
    end
  end

  ##

  defp source({:path, path}) do
    {:ok, String.to_charlist(path)}
  end

  defp source({:binary, binary}) do
    {:ok, binary}
  end

  defp collect_files(entries) do
    entries
    |> Enum.reduce([], fn entry, acc ->
      case entry do
        {:zip_file, path, _, _, _, _} ->
          [to_string(path) | acc]

        _ ->
          acc
      end
    end)
    |> Enum.sort()
  end

  def extract_options(file) do
    [{:file_list, [String.to_charlist(file)]}, :memory]
  end

  defp translate_zip_error({:error, :enoent}) do
    {:error, "file not found"}
  end

  defp translate_zip_error({:error, code})
       when code in [:einval, :bad_eocd, :bad_central_directory, :eisdir] do
    {:error, "invalid zip file"}
  end

  defp translate_zip_error({:error, {:bad_crc, _file}}) do
    {:error, "invalid zip file"}
  end

  # e.g. encrypted files, which `:zip` doesn't support
  defp translate_zip_error({:error, {:EXIT, {:data_error, _stacktrace}}}) do
    {:error, "invalid zip file"}
  end
end

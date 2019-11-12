defmodule XlsxReader.Unzip do
  @type zip_handle :: {:path, String.t()} | {:binary, binary()}

  def list(zip_handle) do
    with {:ok, zip} <- source(zip_handle),
         {:ok, entries} <- :zip.list_dir(zip) do
      {:ok, collect_files(entries)}
    end
  end

  def extract(zip_handle, file) do
    with {:ok, zip} <- source(zip_handle),
         {:ok, [{_, contents}]} <- :zip.extract(zip, extract_options(file)) do
      {:ok, contents}
    else
      {:ok, []} ->
        {:error, "file #{inspect(file)} not found"}

      error ->
        error
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
end

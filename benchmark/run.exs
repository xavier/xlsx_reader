defmodule Benchmark do
  def files(), do: Path.wildcard("benchmark/files/*.xlsx")

  def run(files) do
    timeout = 60_000

    Benchee.run(
      %{
        "XlsxReader.sheets/2" => fn package ->
          XlsxReader.sheets(package)
        end,
        "XlsxReader.async_sheets/3 - ordered" => fn package ->
          XlsxReader.async_sheets(package, [], timeout: timeout)
        end,
        "XlsxReader.async_sheets/3 - unordered" => fn package ->
          XlsxReader.async_sheets(package, [], ordered: false, timeout: timeout)
        end
      },
      inputs: for(file <- files, do: {Path.basename(file), file}),
      before_scenario: fn file ->
        IO.puts("Opening #{file}...")
        {:ok, package} = XlsxReader.open(file)
        package
      end,
      time: 10,
      memory_time: 2,
      print: [configuration: false],
      save: [path: "benchmark/output/save.benchee", tag: "previous"],
      load: "benchmark/output/save.benchee"
    )
  end
end

case Benchmark.files() do
  [] ->
    IO.puts("Benchmarking data is missing. Please first run: mix run benchmark/init.exs")
    System.halt(1)

  files ->
    Benchmark.run(files)
end

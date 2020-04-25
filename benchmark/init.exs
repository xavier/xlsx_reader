defmodule Generator do
  def generate_workbook(sheets, rows, cols, length) do
    %Elixlsx.Workbook{
      sheets: Enum.map(1..sheets, &generate_sheet(&1, rows, cols, length))
    }
  end

  def generate_sheet(index, rows, cols, length) do
    %Elixlsx.Sheet{
      name: "Sheet #{index}",
      rows:
        for _row <- 1..rows do
          Enum.map(1..cols, fn _ -> generate_string(length) end)
        end
    }
  end

  def generate_string(length) do
    length
    |> :crypto.strong_rand_bytes()
    |> Base.encode32()
    |> binary_part(0, length)
  end
end

benchmarks = [
  {"benchmark/files/01_small.xlsx", 4, 10, 8, 10},
  {"benchmark/files/02_medium.xlsx", 8, 1024, 16, 32},
  {"benchmark/files/03_large.xlsx", 16, 4096, 32, 48}
]

Enum.each(benchmarks, fn {file, sheets, rows, cols, length} ->
  IO.puts("Generating #{file}…")
  Elixlsx.write_to(Generator.generate_workbook(sheets, rows, cols, length), file)
end)

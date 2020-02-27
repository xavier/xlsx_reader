defmodule XlsxReader.CellReference do
  @moduledoc false

  @spec parse(String.t()) :: {pos_integer(), pos_integer()} | :error
  def parse(reference) when not is_nil(reference) do
    case Regex.run(~r/\A([A-Z]+)(\d+)\z/, reference, capture: :all_but_first) do
      [letters, digits] ->
        {column_number(letters), String.to_integer(digits)}

      _ ->
        :error
    end
  end

  defp column_number(letters) do
    letters
    |> String.to_charlist()
    |> Enum.reduce(0, fn character_code, column_number ->
      column_number * 26 + (character_code - ?A) + 1
    end)
  end
end

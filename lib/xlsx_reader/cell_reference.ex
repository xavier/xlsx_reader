defmodule XlsxReader.CellReference do
  @moduledoc false

  @spec parse(String.t()) :: {pos_integer(), pos_integer()} | :error
  # Reference starts with an uppercase letter: begin scanning the column letters.
  def parse(<<c, _::binary>> = reference) when c >= ?A and c <= ?Z do
    parse_letters(reference, 0)
  end

  # Anything else (no leading letter, non-binary, …) is not a valid reference.
  def parse(_), do: :error

  # Consume one more A–Z letter into the column accumulator (base-26, A=1).
  defp parse_letters(<<c, rest::binary>>, column) when c >= ?A and c <= ?Z do
    parse_letters(rest, column * 26 + (c - ?A) + 1)
  end

  # First digit reached after at least one letter: switch to scanning the row digits.
  defp parse_letters(<<c, _::binary>> = rest, column) when c >= ?0 and c <= ?9 and column > 0 do
    parse_digits(rest, 0)
    |> finalize(column)
  end

  # Unexpected character or end of input before any digit.
  defp parse_letters(_, _), do: :error

  # Consume one more 0–9 digit into the row accumulator.
  defp parse_digits(<<c, rest::binary>>, row) when c >= ?0 and c <= ?9 do
    parse_digits(rest, row * 10 + (c - ?0))
  end

  # End of input after at least one digit: the row is complete.
  defp parse_digits("", row) when row > 0, do: row

  # Trailing non-digit, or no digits at all: invalid.
  defp parse_digits(_, _), do: :error

  # parse_digits returned a row integer: combine with the column to produce {col, row}.
  defp finalize(row, column) when is_integer(row), do: {column, row}

  # parse_digits returned :error: propagate it.
  defp finalize(_, _), do: :error
end

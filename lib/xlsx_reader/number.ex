defprotocol XlsxReader.Number do
  @moduledoc false
  def multiply(number, factor)
end

defimpl XlsxReader.Number, for: Integer do
  def multiply(number, factor), do: number * factor
end

defimpl XlsxReader.Number, for: Float do
  def multiply(number, factor), do: number * factor
end

defimpl XlsxReader.Number, for: Decimal do
  def multiply(number, factor), do: Decimal.mult(number, factor)
end

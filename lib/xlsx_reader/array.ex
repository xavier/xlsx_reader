defmodule XlsxReader.Array do
  @moduledoc false

  @type t :: :array.array()
  @type t(type) :: :array.array(type)

  def from_list(list) do
    :array.from_list(list, nil)
  end

  def insert(array, index, value) do
    :array.set(index, value, array)
  end

  def lookup(array, index, default \\ nil) do
    case :array.get(index, array) do
      :undefined ->
        default

      value ->
        value
    end
  end
end

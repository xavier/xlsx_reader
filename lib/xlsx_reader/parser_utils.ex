defmodule XlsxReader.ParserUtils do
  @moduledoc """

  Utility functions used by the XML parser modules

  """

  @type xml_attribute :: {String.t(), String.t()}
  @type xml_attributes :: [xml_attribute]

  @doc """

  Get value of attribute by name

  ## Examples

      iex> XlsxReader.ParserUtils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "y"}], "a")
      "x"

      iex> XlsxReader.ParserUtils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "y"}], "b")
      "y"

      iex> XlsxReader.ParserUtils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "z"}], "c")
      "z"

      iex> XlsxReader.ParserUtils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "z"}], "d")
      nil

      iex> XlsxReader.ParserUtils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "z"}], "d", "default")
      "default"

  """
  @spec get_attribute(xml_attributes(), String.t(), nil | String.t()) ::
          nil | String.t()
  def get_attribute(attributes, name, default \\ nil)
  def get_attribute([], _name, default), do: default
  def get_attribute([{name, value} | _], name, _default), do: value
  def get_attribute([_ | rest], name, default), do: get_attribute(rest, name, default)

  @doc """

  Extracts XML attributes into to map based on the given mapping

  ## Examples

      iex> XlsxReader.ParserUtils.map_attributes([{"a", "x"}, {"b", "y"}], %{"a" => :foo, "b" => :bar, "c" => :baz})
      %{foo: "x", bar: "y"}

  """

  @spec map_attributes(xml_attributes(), map(), map()) :: map()
  def map_attributes(attributes, mapping, initial \\ %{}) do
    Enum.reduce(attributes, initial, fn {name, value}, acc ->
      case Map.fetch(mapping, name) do
        {:ok, key} ->
          Map.put(acc, key, value)

        :error ->
          acc
      end
    end)
  end
end

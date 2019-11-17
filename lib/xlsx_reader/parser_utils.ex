defmodule XlsxReader.ParserUtils do
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
  @spec get_attribute([{String.t(), String.t()}], String.t(), nil | String.t()) ::
          nil | String.t()
  def get_attribute(attributes, name, default \\ nil)
  def get_attribute([], _name, default), do: default
  def get_attribute([{name, value} | _], name, _default), do: value
  def get_attribute([_ | rest], name, default), do: get_attribute(rest, name, default)
end

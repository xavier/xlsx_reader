defmodule XlsxReader.Parsers.Utils do
  @moduledoc false

  # Utility functions used by the XML parser modules

  @type xml_attribute :: {String.t(), String.t()}
  @type xml_attributes :: [xml_attribute]

  @doc """

  Get value of attribute by name

  ## Examples

      iex> XlsxReader.Parsers.Utils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "y"}], "a")
      "x"

      iex> XlsxReader.Parsers.Utils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "y"}], "b")
      "y"

      iex> XlsxReader.Parsers.Utils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "z"}], "c")
      "z"

      iex> XlsxReader.Parsers.Utils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "z"}], "d")
      nil

      iex> XlsxReader.Parsers.Utils.get_attribute([{"a", "x"}, {"b", "y"}, {"c", "z"}], "d", "default")
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

      iex> XlsxReader.Parsers.Utils.map_attributes([{"a", "x"}, {"b", "y"}], %{"a" => :foo, "b" => :bar, "c" => :baz})
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

  @doc """

  Returns an UTF-8 binary which is the only character encoding supported by the XML parser.

  Converts to UTF-8 from UTF-16BE/LE if a BOM is detected.

  ## Examples

      iex> XlsxReader.Parsers.Utils.ensure_utf8("UTF-8")
      {:ok, "UTF-8"}

      iex> XlsxReader.Parsers.Utils.ensure_utf8(<<0xff, 0xfe, 0x55, 0x00, 0x54, 0x00, 0x46, 0x00, 0x2d, 0x00, 0x31, 0x00, 0x36, 0x00, 0x4c, 0x00, 0x45, 0x00>>)
      {:ok, "UTF-16LE"}

      iex> XlsxReader.Parsers.Utils.ensure_utf8(<<0xfe, 0xff, 0x00, 0x55, 0x00, 0x54, 0x00, 0x46, 0x00, 0x2d, 0x00, 0x31, 0x00, 0x36, 0x00, 0x42, 0x00, 0x45>>)
      {:ok, "UTF-16BE"}

      iex> XlsxReader.Parsers.Utils.ensure_utf8(<<0xff, 0xfe, 0x00>>)
      {:error, "incomplete UTF-16 binary"}

  """
  @spec ensure_utf8(binary()) :: {:ok, String.t()} | {:error, String.t()}
  def ensure_utf8(<<0xFF, 0xFE, rest::binary>>),
    do: convert_utf16_to_utf8(rest, :little)

  def ensure_utf8(<<0xFE, 0xFF, rest::binary>>),
    do: convert_utf16_to_utf8(rest, :big)

  def ensure_utf8(utf8), do: {:ok, utf8}

  defp convert_utf16_to_utf8(utf16, endianess) do
    case :unicode.characters_to_binary(utf16, {:utf16, endianess}) do
      utf8 when is_binary(utf8) ->
        {:ok, utf8}

      {:error, _, _} ->
        {:error, "error converting UTF-16 binary to UTF-8"}

      {:incomplete, _, _} ->
        {:error, "incomplete UTF-16 binary"}
    end
  end
end

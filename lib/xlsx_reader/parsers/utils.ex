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

      iex> XlsxReader.Parsers.Utils.ensure_utf8("\uFEFFUTF-8 with BOM")
      {:ok, "UTF-8 with BOM"}

      iex> XlsxReader.Parsers.Utils.ensure_utf8(<<0xff, 0xfe, 0x55, 0x00, 0x54, 0x00, 0x46, 0x00, 0x2d, 0x00, 0x31, 0x00, 0x36, 0x00, 0x4c, 0x00, 0x45, 0x00>>)
      {:ok, "UTF-16LE"}

      iex> XlsxReader.Parsers.Utils.ensure_utf8(<<0xfe, 0xff, 0x00, 0x55, 0x00, 0x54, 0x00, 0x46, 0x00, 0x2d, 0x00, 0x31, 0x00, 0x36, 0x00, 0x42, 0x00, 0x45>>)
      {:ok, "UTF-16BE"}

      iex> XlsxReader.Parsers.Utils.ensure_utf8(<<0xff, 0xfe, 0x00>>)
      {:error, "incomplete UTF-16LE binary"}

  """
  @spec ensure_utf8(binary()) :: {:ok, String.t()} | {:error, String.t()}
  def ensure_utf8(string) do
    case :unicode.bom_to_encoding(string) do
      {:latin1, 0} ->
        # No BOM found, assumes UTF-8
        {:ok, string}

      {:utf8, bom_length} ->
        # BOM found with UTF-8 encoding
        {:ok, strip_bom(string, bom_length)}

      {encoding, bom_length} ->
        # BOM found with UTF-16/32 encoding given as an {encoding, endianess} tuple
        string |> strip_bom(bom_length) |> convert_to_utf8(encoding)
    end
  end

  defp strip_bom(string, bom_length) do
    :binary.part(string, {bom_length, byte_size(string) - bom_length})
  end

  defp convert_to_utf8(string, encoding) do
    case :unicode.characters_to_binary(string, encoding) do
      utf8 when is_binary(utf8) ->
        {:ok, utf8}

      {:error, _, _} ->
        {:error, "error converting #{format_encoding(encoding)} binary to UTF-8"}

      {:incomplete, _, _} ->
        {:error, "incomplete #{format_encoding(encoding)} binary"}
    end
  end

  defp format_encoding({:utf16, endianess}), do: "UTF-16#{format_endianess(endianess)}"
  defp format_encoding({:utf32, endianess}), do: "UTF-32#{format_endianess(endianess)}"

  def format_endianess(:big), do: "BE"
  def format_endianess(:little), do: "LE"
end

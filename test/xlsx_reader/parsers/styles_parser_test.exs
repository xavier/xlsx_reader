defmodule XlsxReader.Parsers.StylesParserTest do
  use ExUnit.Case

  alias XlsxReader.Array
  alias XlsxReader.Parsers.StylesParser

  setup do
    {:ok, %{styles_xml: TestFixtures.read!("package/xl/styles.xml")}}
  end

  test "parses styles.xml into an array of style types", %{styles_xml: styles_xml} do
    expected_style_types =
      Array.from_list([
        :number,
        :number,
        :unsupported,
        :number,
        :number,
        :number,
        :number,
        :number,
        :number,
        :number,
        :number,
        :unsupported,
        :unsupported,
        :number,
        :number,
        :number,
        :number,
        :number,
        :date,
        :unsupported,
        :date_time,
        :time,
        :percentage,
        nil,
        :unsupported
      ])

    assert {:ok, ^expected_style_types, _custom_formats} = StylesParser.parse(styles_xml)
  end

  test "supports user-provided custom formats", %{styles_xml: styles_xml} do
    supported_custom_formats = [
      {"[$CHF]0.00", :number}
    ]

    expected_style_types =
      Array.from_list([
        :number,
        :number,
        :unsupported,
        :number,
        :number,
        :number,
        :number,
        :number,
        :number,
        :number,
        :number,
        :unsupported,
        :unsupported,
        :number,
        :number,
        :number,
        :number,
        :number,
        :date,
        :unsupported,
        :date_time,
        :time,
        :percentage,
        :number,
        :unsupported
      ])

    assert {:ok, ^expected_style_types, _custom_formats} =
             StylesParser.parse(styles_xml, supported_custom_formats)
  end
end

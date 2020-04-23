defmodule XlsxReader.Parsers.StylesParserTest do
  use ExUnit.Case

  alias XlsxReader.Array
  alias XlsxReader.Parsers.StylesParser

  test "parses styles.xml into an array of style types" do
    styles_xml = TestFixtures.read!("package/xl/styles.xml")

    expected =
      Array.from_list([
        :string,
        :string,
        :unsupported,
        :string,
        :string,
        :string,
        :string,
        :string,
        :string,
        :string,
        :string,
        :unsupported,
        :unsupported,
        :string,
        :string,
        :string,
        :string,
        :string,
        :date,
        :unsupported,
        :date_time,
        :time,
        :percentage,
        "[$CHF]\" \"0.00",
        :unsupported
      ])

    assert {:ok, expected} == StylesParser.parse(styles_xml)
  end
end

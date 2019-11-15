defmodule XlsxReader.StylesParserTest do
  use ExUnit.Case

  alias XlsxReader.StylesParser

  test "parses sharedStrings.xml" do
    styles_xml = TestFixtures.read!("package/xl/styles.xml")

    expected = %{
      custom_formats: %{
        "0" => "General",
        "59" => "dd/mm/yyyy",
        "60" => "dd/mm/yyyy hh:mm",
        "61" => "hh:mm",
        "62" => "0.0%",
        "63" => "[$CHF] 0.00"
      },
      style_types: [
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
        "[$CHF] 0.00",
        :unsupported,
        :percentage
      ]
    }

    assert {:ok, expected} == StylesParser.parse(styles_xml)
  end
end

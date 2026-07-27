defmodule XlsxReader.Parsers.StylesParserTest do
  use ExUnit.Case

  alias XlsxReader.Array
  alias XlsxReader.Parsers.StylesParser

  setup do
    {:ok, %{styles_xml: TestFixtures.read!("package/xl/styles.xml")}}
  end

  defp style_types_from_num_fmts(num_fmts) do
    num_fmts
    |> Enum.map(&%{num_fmt: &1, fill: nil})
    |> Array.from_list()
  end

  test "parses styles.xml into an array of style types", %{styles_xml: styles_xml} do
    expected_style_types =
      style_types_from_num_fmts([
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
      style_types_from_num_fmts([
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

  test "collects fill background colors" do
    styles_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <styleSheet>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill>
          <patternFill patternType="solid">
            <fgColor rgb="FFFF0000"/>
            <bgColor indexed="64"/>
          </patternFill>
        </fill>
        <fill>
          <patternFill patternType="solid">
            <bgColor rgb="FF00FF00"/>
          </patternFill>
        </fill>
      </fills>
      <cellXfs count="3">
        <xf numFmtId="0"/>
        <xf numFmtId="0" fillId="1"/>
        <xf numFmtId="0" fillId="2"/>
      </cellXfs>
    </styleSheet>
    """

    expected_style_types =
      Array.from_list([
        %{num_fmt: :number, fill: nil},
        # `bgColor indexed="64"` resolves to the pattern's foreground color
        %{num_fmt: :number, fill: %{fg_color: "FFFF0000", bg_color: "FFFF0000"}},
        %{num_fmt: :number, fill: %{bg_color: "FF00FF00"}}
      ])

    assert {:ok, ^expected_style_types, _custom_formats} = StylesParser.parse(styles_xml)
  end
end

defmodule XlsxReader.Parsers.StylesParserTest do
  use ExUnit.Case

  alias XlsxReader.Parsers.StylesParser

  setup do
    {:ok, %{styles_xml: TestFixtures.read!("package/xl/styles.xml")}}
  end

  test "parses styles.xml into an array of style types", %{styles_xml: styles_xml} do
    assert {:ok, style_types, _custom_formats} = StylesParser.parse(styles_xml)

    assert num_fmts(style_types) == [
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
           ]
  end

  test "supports user-provided custom formats", %{styles_xml: styles_xml} do
    supported_custom_formats = [
      {"[$CHF]0.00", :number}
    ]

    assert {:ok, style_types, _custom_formats} =
             StylesParser.parse(styles_xml, supported_custom_formats)

    assert num_fmts(style_types) == [
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
           ]
  end

  test "parses the cell background colors", %{styles_xml: styles_xml} do
    assert {:ok, style_types, _custom_formats} = StylesParser.parse(styles_xml)

    # The fills of this fixture are solid patterns whose foreground color
    # refers to the palette redefined by the `indexedColors` element.
    assert bg_colors(style_types) == [
             nil,
             nil,
             "ffbdc0bf",
             "ffdbdbdb",
             "ffffffff",
             "ffffffff",
             "ffdbdbdb",
             "ffffffff",
             "ffffffff",
             nil,
             "ffbdc0bf",
             "ffdbdbdb",
             "ffffffff",
             "ffffffff",
             "ffdbdbdb",
             "ffffffff",
             "ffffffff",
             nil,
             "ffffffff",
             "ffdbdbdb",
             "ffffffff",
             "ffffffff",
             "ffffffff",
             "ffffffff",
             "ffffffff"
           ]
  end

  test "resolves the background color according to the fill pattern type" do
    styles_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <styleSheet>
      <fills count="5">
        <fill><patternFill patternType="none"/></fill>
        <fill>
          <patternFill patternType="solid">
            <fgColor rgb="FFFF0000"/>
            <bgColor indexed="64"/>
          </patternFill>
        </fill>
        <fill>
          <patternFill patternType="darkUp">
            <fgColor rgb="FF000000"/>
            <bgColor rgb="FF00FF00"/>
          </patternFill>
        </fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill>
          <patternFill patternType="solid">
            <bgColor rgb="FF0000FF"/>
          </patternFill>
        </fill>
        <fill>
          <gradientFill degree="90">
            <stop position="0"><color rgb="FF0000FF"/></stop>
            <stop position="1"><color rgb="FFFFFFFF"/></stop>
          </gradientFill>
        </fill>
      </fills>
      <cellXfs count="7">
        <xf numFmtId="0"/>
        <xf numFmtId="0" fillId="0"/>
        <xf numFmtId="0" fillId="1"/>
        <xf numFmtId="0" fillId="2"/>
        <xf numFmtId="0" fillId="3"/>
        <xf numFmtId="0" fillId="4"/>
        <xf numFmtId="0" fillId="5"/>
      </cellXfs>
    </styleSheet>
    """

    assert {:ok, style_types, _custom_formats} = StylesParser.parse(styles_xml)

    assert bg_colors(style_types) == [
             # no fill
             nil,
             # patternType="none"
             nil,
             # solid pattern: the foreground color is the visible background
             "FFFF0000",
             # other patterns: the background the pattern is drawn over
             "FF00FF00",
             # pattern without color
             nil,
             # solid pattern without foreground color: bgColor is unused
             nil,
             # gradient fills are not supported
             nil
           ]
  end

  test "ignores fills which paint nothing" do
    styles_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <styleSheet>
      <fills count="4">
        <fill/>
        <fill><patternFill/></fill>
        <fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill>
        <fill><patternFill patternType="none"><bgColor rgb="FF00FF00"/></patternFill></fill>
      </fills>
      <cellXfs count="5">
        <xf numFmtId="0" fillId="0"/>
        <xf numFmtId="0" fillId="1"/>
        <xf numFmtId="0" fillId="2"/>
        <xf numFmtId="0" fillId="3"/>
        <xf numFmtId="0" fillId="9"/>
      </cellXfs>
    </styleSheet>
    """

    assert {:ok, style_types, _custom_formats} = StylesParser.parse(styles_xml)

    # A fill only paints when it declares a pattern type, whatever colors it
    # carries. The last format refers to a fill which does not exist.
    assert bg_colors(style_types) == [nil, nil, nil, nil, nil]
  end

  test "resolves indexed colors against the default palette" do
    styles_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <styleSheet>
      <fills count="4">
        <fill><patternFill patternType="solid"><fgColor indexed="10"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor indexed="22"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor auto="1"/></patternFill></fill>
      </fills>
      <cellXfs count="4">
        <xf numFmtId="0" fillId="0"/>
        <xf numFmtId="0" fillId="1"/>
        <xf numFmtId="0" fillId="2"/>
        <xf numFmtId="0" fillId="3"/>
      </cellXfs>
    </styleSheet>
    """

    assert {:ok, style_types, _custom_formats} = StylesParser.parse(styles_xml)

    assert bg_colors(style_types) == [
             "00FF0000",
             "00C0C0C0",
             # index 64 is the system foreground color
             nil,
             # automatic colors are unknown
             nil
           ]
  end

  test "resolves indexed colors against the palette redefined by the stylesheet" do
    styles_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <styleSheet>
      <fills count="3">
        <fill><patternFill patternType="solid"><fgColor indexed="1"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor indexed="2"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor indexed="22"/></patternFill></fill>
      </fills>
      <cellXfs count="3">
        <xf numFmtId="0" fillId="0"/>
        <xf numFmtId="0" fillId="1"/>
        <xf numFmtId="0" fillId="2"/>
      </cellXfs>
      <colors>
        <indexedColors>
          <rgbColor rgb="00000000"/>
          <rgbColor rgb="00112233"/>
          <rgbColor rgb="00445566"/>
        </indexedColors>
      </colors>
    </styleSheet>
    """

    assert {:ok, style_types, _custom_formats} = StylesParser.parse(styles_xml)

    assert bg_colors(style_types) == [
             "00112233",
             "00445566",
             # indexes left out of the redefined palette fall back to the default
             "00C0C0C0"
           ]
  end

  test "ignores the fills of the differential formats" do
    styles_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <styleSheet>
      <fills count="1">
        <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
      </fills>
      <cellXfs count="1">
        <xf numFmtId="0" fillId="0"/>
      </cellXfs>
      <dxfs count="1">
        <dxf>
          <fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill>
        </dxf>
      </dxfs>
    </styleSheet>
    """

    assert {:ok, style_types, _custom_formats} = StylesParser.parse(styles_xml)

    assert bg_colors(style_types) == ["FFFF0000"]
  end

  defp num_fmts(style_types),
    do: style_types |> :array.to_list() |> Enum.map(& &1.num_fmt)

  defp bg_colors(style_types),
    do: style_types |> :array.to_list() |> Enum.map(& &1.bg_color)
end

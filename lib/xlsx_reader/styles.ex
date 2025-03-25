defmodule XlsxReader.Styles do
  @moduledoc """
  Utility functions and types to deal with cell styles
  """

  @typedoc """
  Knwon cell styles for which type conversion is supported
  """
  @type known_style_type ::
          :string | :number | :percentage | :date | :time | :date_time | :unsupported
  @typedoc """
  Cell styles for which type conversion is supported
  """
  @type style_type :: known_style_type() | String.t()
  @type style_types :: XlsxReader.Array.t(style_type())
  @type custom_formats :: %{optional(String.t()) => String.t()}
  @typedoc """
  Matches the text representation of a cell value
  """
  @type custom_format_matcher :: String.t() | Regex.t()
  @type supported_custom_formats :: [{custom_format_matcher(), known_style_type()}]

  @known_styles %{
    # General
    "0" => :number,
    # 0
    "1" => :number,
    # 0.00
    "2" => :number,
    # #,##0
    "3" => :number,
    # #,##0.00
    "4" => :number,
    # $#,##0_);($#,##0)
    "5" => :unsupported,
    # $#,##0_);[Red]($#,##0)
    "6" => :unsupported,
    # $#,##0.00_);($#,##0.00)
    "7" => :unsupported,
    # $#,##0.00_);[Red]($#,##0.00)
    "8" => :unsupported,
    # 0%
    "9" => :percentage,
    # 0.00%
    "10" => :percentage,
    # 0.00E+00
    "11" => :number,
    # # ?/?
    "12" => :unsupported,
    # # ??/??
    "13" => :unsupported,
    # mm-dd-yy
    "14" => :date,
    # d-mmm-yy
    "15" => :date,
    # d-mmm
    "16" => :date,
    # mmm-yy
    "17" => :date,
    # h:mm AM/PM
    "18" => :time,
    # h:mm:ss AM/PM
    "19" => :time,
    # h:mm
    "20" => :time,
    # h:mm:ss
    "21" => :time,
    # m/d/yy h:mm
    "22" => :date_time,
    # #,##0 ;(#,##0)
    "37" => :unsupported,
    # #,##0 ;[Red](#,##0)
    "38" => :unsupported,
    # #,##0.00;(#,##0.00)
    "39" => :unsupported,
    # #,##0.00;[Red](#,##0.00)
    "40" => :unsupported,
    # mm:ss
    "45" => :time,
    # [h]:mm:ss
    "46" => :time,
    # mmss.0
    "47" => :time,
    # ##0.0E+0
    "48" => :number,
    # @
    "49" => :unsupported
  }

  @default_supported_custom_formats [
    {"0.0%", :percentage},
    {~r/\Add?\/mm?\/yy(?:yy)\z/, :date},
    {~r/\Add?\/mm?\/yy(?:yy) hh?:mm?\z/, :date_time},
    {~r/\Ay+(\\\/|-)m+(\\\/|-)d+\z/i, :date},
    {~r/\Ayyyy-mm-dd[T\s]hh?:mm:ssZ?\z/, :date_time},
    {"m/d/yyyy", :date},
    {"m/d/yyyy h:mm", :date_time},
    {"hh:mm", :time}
  ]

  @doc """
  Guesses the type of a cell based on its style.

  The type is:

  1. looked-up from a list of "standard" styles, or
  2. guessed from a list of default supported custom formats, or
  3. guessed from a list of user-provided supported custom formats.

  If no type could be guessed, returns `nil`.

  """
  @spec get_style_type(String.t(), custom_formats(), supported_custom_formats()) ::
          style_type() | nil
  def get_style_type(num_fmt_id, custom_formats \\ %{}, supported_custom_formats \\ []) do
    get_known_style(num_fmt_id) ||
      get_custom_style(num_fmt_id, custom_formats, supported_custom_formats)
  end

  defp get_known_style(num_fmt_id),
    do: Map.get(@known_styles, num_fmt_id)

  defp get_custom_style(num_fmt_id, custom_formats, supported_custom_formats) do
    get_style_type_from_custom_format(
      num_fmt_id,
      custom_formats,
      @default_supported_custom_formats
    ) ||
      get_style_type_from_custom_format(
        num_fmt_id,
        custom_formats,
        supported_custom_formats
      )
  end

  defp get_style_type_from_custom_format(num_fmt_id, custom_formats, supported_custom_format) do
    custom_formats
    |> Map.get(num_fmt_id)
    |> custom_format_to_style_type(supported_custom_format)
  end

  defp custom_format_to_style_type(nil, _), do: nil
  defp custom_format_to_style_type(_custom_format, []), do: nil

  defp custom_format_to_style_type(custom_format, [{%Regex{} = regex, style_type} | others]) do
    if Regex.match?(regex, custom_format),
      do: style_type,
      else: custom_format_to_style_type(custom_format, others)
  end

  defp custom_format_to_style_type(custom_format, [{custom_format, style_type} | _others]),
    do: style_type

  defp custom_format_to_style_type(custom_format, [_ | others]),
    do: custom_format_to_style_type(custom_format, others)
end

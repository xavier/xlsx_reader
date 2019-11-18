defmodule XlsxReader.Conversion do
  @moduledoc """

  Conversion of cell values to Elixir types

  """

  @type date_system :: 1900 | 1904

  @doc """

  Converts a string to an integer or float

  ## Examples

      iex> XlsxReader.Conversion.to_number("123")
      {:ok, 123}

      iex> XlsxReader.Conversion.to_number("123.0")
      {:ok, 123.0}

      iex> XlsxReader.Conversion.to_number("-123.45")
      {:ok, -123.45}

      iex> XlsxReader.Conversion.to_number("0.12345e3")
      {:ok, 123.45}

      iex> XlsxReader.Conversion.to_number("0.12345E3")
      {:ok, 123.45}

  """
  @spec to_number(String.t()) :: {:ok, integer()} | {:ok, float()} | :error
  def to_number(string) do
    if String.match?(string, ~r/[\.eE]/),
      do: to_float(string),
      else: to_integer(string)
  end

  @doc """

  Converts a string to float

  ## Examples

      iex> XlsxReader.Conversion.to_float("123")
      {:ok, 123.0}

      iex> XlsxReader.Conversion.to_float("-123.45")
      {:ok, -123.45}

      iex> XlsxReader.Conversion.to_float("0.12345e3")
      {:ok, 123.45}

      iex> XlsxReader.Conversion.to_float("0.12345E3")
      {:ok, 123.45}

  """
  @spec to_float(String.t()) :: {:ok, float()} | :error
  def to_float(string) do
    case Float.parse(string) do
      {number, ""} ->
        {:ok, number}

      _ ->
        :error
    end
  end

  @doc """

  Converts a string to an integer

  ## Examples

      iex> XlsxReader.Conversion.to_integer("123")
      {:ok, 123}

      iex> XlsxReader.Conversion.to_integer("-123")
      {:ok, -123}

      iex> XlsxReader.Conversion.to_integer("123.45")
      :error

  """
  @spec to_integer(String.t()) :: {:ok, integer()} | :error
  def to_integer(string) do
    case Integer.parse(string) do
      {number, ""} ->
        {:ok, number}

      _ ->
        :error
    end
  end

  # This is why we can't have nice things: http://www.cpearson.com/excel/datetime.htm
  @base_date_system_1900 ~D[1899-12-30]
  @base_date_system_1904 ~D[1904-01-01]

  @doc """
  Returns the base date for the given date system
  """
  @spec base_date(date_system()) :: Date.t()
  def base_date(1900), do: @base_date_system_1900
  def base_date(1904), do: @base_date_system_1904

  @doc """

  Converts a serial date to a `Date`

  ## Examples

      iex> XlsxReader.Conversion.to_date("40396")
      {:ok, ~D[2010-08-06]}

      iex> XlsxReader.Conversion.to_date("43783")
      {:ok, ~D[2019-11-14]}

      iex> XlsxReader.Conversion.to_date("1", ~D[1999-12-31])
      {:ok, ~D[2000-01-01]}

      iex> XlsxReader.Conversion.to_date("-1", ~D[1999-12-31])
      :error

  """
  @spec to_date(String.t(), Date.t()) :: {:ok, Date.t()} | :error
  def to_date(string, base_date \\ @base_date_system_1900) do
    case split_serial_date(string) do
      {:ok, days, _fraction_of_24} when days > 0.0 ->
        {:ok, Date.add(base_date, days)}

      {:ok, _days, _fraction_of_24} ->
        :error

      error ->
        error
    end
  end

  @doc """

  Converts a serial date to a `NaiveDateTime`

  ## Examples

      iex> XlsxReader.Conversion.to_date_time("43783.0")
      {:ok, ~N[2019-11-14 00:00:00]}

      iex> XlsxReader.Conversion.to_date_time("43783.760243055556")
      {:ok, ~N[2019-11-14 18:14:45]}

      iex> XlsxReader.Conversion.to_date_time("1.760243055556", ~D[1999-12-31])
      {:ok, ~N[2000-01-01 18:14:45]}

      iex> XlsxReader.Conversion.to_date_time("-30.760243055556", ~D[1999-12-31])
      :error

  """
  @spec to_date_time(String.t(), Date.t()) :: {:ok, NaiveDateTime.t()} | :error
  def to_date_time(string, base_date \\ @base_date_system_1900) do
    with {:ok, days, fraction_of_24} when days > 0.0 <- split_serial_date(string),
         date <- Date.add(base_date, days),
         {:ok, time} <- fraction_of_24_to_time(fraction_of_24) do
      NaiveDateTime.new(date, time)
    else
      {:ok, _, _} ->
        :error

      error ->
        error
    end
  end

  ## Private

  @spec split_serial_date(String.t()) :: {:ok, integer(), float()} | :error
  defp split_serial_date(string) do
    with {:ok, value} <- to_float(string) do
      days = Float.floor(value)
      {:ok, trunc(days), value - days}
    end
  end

  @seconds_per_day 60 * 60 * 24

  @spec fraction_of_24_to_time(float()) :: {:ok, Time.t()}
  defp fraction_of_24_to_time(fraction_of_24) do
    seconds = round(fraction_of_24 * @seconds_per_day)

    Time.new(
      seconds |> div(3600),
      seconds |> div(60) |> rem(60),
      seconds |> rem(60)
    )
  end
end

defmodule CustomDatesTest do
  @moduledoc """
  Ensure that various custom date and
  datetime formats are parsed into
  Date or NaiveDateTime as appropriate.
  """

  use ExUnit.Case

  test "custom_dates.xlsx" do
    assert {:ok, package} = XlsxReader.open(TestFixtures.path("custom_dates.xlsx"))

    assert ["Sheet1"] = XlsxReader.sheet_names(package)

    assert {:ok,
            [
              ["ISO8601 Date", ~D[2020-05-01]],
              ["ISO8601 Datetime", ~N[2020-05-01 12:45:59]],
              ["US Date", ~D[2020-05-01]],
              ["US Date", ~D[2020-12-31]],
              ["US Datetime", ~N[2020-05-01 01:23:00]],
              ["US Datetime", ~N[2020-05-01 12:23:00]]
            ]} = XlsxReader.sheet(package, "Sheet1")
  end
end

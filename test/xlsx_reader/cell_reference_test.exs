defmodule XlsxReader.CellReferenceTest do
  use ExUnit.Case

  alias XlsxReader.CellReference

  describe ".parse/1" do
    test "returns a {col, row} tuple" do
      assert {1, 1} == CellReference.parse("A1")
      assert {1, 42} == CellReference.parse("A42")
      assert {2, 1} == CellReference.parse("B1")
      assert {26, 3} == CellReference.parse("Z3")
      assert {26 * 1 + 1, 123} == CellReference.parse("AA123")
      assert {26 * 1 + 26, 123} == CellReference.parse("AZ123")
      assert {26 * 2 + 1, 123} == CellReference.parse("BA123")
      assert {26 * 26 + 1, 123} == CellReference.parse("ZA123")
      assert {26 * 26 + 26, 123} == CellReference.parse("ZZ123")
      assert {26 * 26 * 1 + 26 * 1 + 1, 123_456} == CellReference.parse("AAA123456")
      assert {26 * 26 * 1 + 26 * 1 + 26, 123_456} == CellReference.parse("AAZ123456")
      assert {26 * 26 * 1 + 26 * 2 + 1, 123_456} == CellReference.parse("ABA123456")
      assert {26 * 26 * 2 + 26 * 26 + 1, 123_456} == CellReference.parse("BZA123456")
    end

    test "returns error if the reference is invalid" do
      assert :error == CellReference.parse("1A")
      assert :error == CellReference.parse("$A1")
      assert :error == CellReference.parse("A$1")
      assert :error == CellReference.parse("$A$1")
      assert :error == CellReference.parse("bogus")
    end
  end
end

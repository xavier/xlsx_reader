defmodule XlsxReader.StylesTest do
  use ExUnit.Case

  alias XlsxReader.Styles

  describe "get_style_type/2" do
    test "some known styles" do
      assert :number == Styles.get_style_type("0", %{})
      assert :number == Styles.get_style_type("1", %{})
      assert :number == Styles.get_style_type("2", %{})
      assert :percentage == Styles.get_style_type("9", %{})
      assert :date == Styles.get_style_type("14", %{})
      assert :time == Styles.get_style_type("18", %{})
      assert :unsupported == Styles.get_style_type("49", %{})
    end

    test "supported custom formats" do
      assert :percentage = Styles.get_style_type("123", %{"123" => "0.0%"})

      # ISO8601 date/time
      assert :date = Styles.get_style_type("123", %{"123" => "yyyy-mm-dd"})
      assert :date_time = Styles.get_style_type("123", %{"123" => "yyyy-mm-dd hh:mm:ss"})
      assert :date_time = Styles.get_style_type("123", %{"123" => "yyyy-mm-ddThh:mm:ssZ"})

      # US date/time
      assert :date = Styles.get_style_type("123", %{"123" => "m/d/yyyy"})
      assert :date_time = Styles.get_style_type("123", %{"123" => "m/d/yyyy h:mm"})

      # Plain time
      assert :time = Styles.get_style_type("123", %{"123" => "hh:mm"})
    end

    test "user-provided supported custom formats" do
      assert :date == Styles.get_style_type("123", %{"123" => "mmm yy"}, [{"mmm yy", :date}])
      assert :date == Styles.get_style_type("123", %{"123" => "mmm yy"}, [{~r/mmm? yy/, :date}])
    end

    test "unknown format" do
      assert nil == Styles.get_style_type("123", %{"456" => "0.0%"})
      assert nil == Styles.get_style_type("123", %{"123" => "bogus"})
      assert nil == Styles.get_style_type("123", %{"123" => "bogus"}, [{"mmm yy", :date}])
    end
  end
end

defmodule XlsxReader.Workbook do
  @moduledoc false

  defstruct sheets: [],
            rels: nil,
            shared_strings: nil,
            style_types: nil
end

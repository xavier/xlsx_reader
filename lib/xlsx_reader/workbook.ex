defmodule XlsxReader.Workbook do
  @moduledoc """

  Workbook structure.

  - `sheets` - list of sheet metadata
  - `rels` - workbook relationships
  - `shared_strings` - list of shared strings
  - `style_types` - List of types indexed by style
  - `base_date` - base date for all serial dates in the workbook

  """

  defstruct sheets: [],
            rels: nil,
            shared_strings: nil,
            style_types: nil,
            base_date: nil

  @typedoc """
  XLSX workbook
  """
  @type t :: %__MODULE__{
          sheets: [XlsxReader.Sheet.t()],
          rels: nil | map(),
          shared_strings: nil | [String.t()],
          style_types: nil | XlsxReader.Styles.style_types(),
          base_date: nil | Date.t()
        }
end

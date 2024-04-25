defmodule XlsxReader.Sheet do
  @moduledoc """

  Worksheet structure.

  This structure contains initially only the information useful to identify and retrieve the sheet actual data.

  - `name` - name of the sheet
  - `rid` - relationship ID used to retrieve the corresponding sheet in the archive
  - `sheet_id` - unique identifier of the sheet withing the workbook (unused by XlsxReader)
  - `data` - `nil` or a list of rows if sheet was loaded

  To access the sheet cells, see `XlsxReader.sheet/3` and `XlsxReader.sheets/2`.

  """

  defstruct [:name, :rid, :sheet_id, :data]

  @typedoc """
  XLSX worksheet metadata
  """
  @type t :: %__MODULE__{
          name: String.t(),
          rid: String.t(),
          sheet_id: String.t(),
          data: nil | [XlsxReader.row()]
        }
end

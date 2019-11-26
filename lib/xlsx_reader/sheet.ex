defmodule XlsxReader.Sheet do
  @moduledoc """

  Worksheet structure.

  - `name` - name of the sheet
  - `rid` - relationship ID used to retrieve the corresponding sheet in the archive
  - `sheet_id` - unused

  """

  defstruct [:name, :rid, :sheet_id]

  @type t :: %__MODULE__{
          name: String.t(),
          rid: String.t(),
          sheet_id: String.t()
        }
end

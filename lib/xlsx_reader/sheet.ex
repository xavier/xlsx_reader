defmodule XlsxReader.Sheet do
  @moduledoc false

  defstruct [:name, :rid, :sheet_id]

  @type t :: %__MODULE__{
          name: String.t(),
          rid: String.t(),
          sheet_id: String.t()
        }
end

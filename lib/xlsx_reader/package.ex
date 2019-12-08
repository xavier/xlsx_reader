defmodule XlsxReader.Package do
  @moduledoc """

  XSLX Package structure.

  This structure is initialized by `XlsxReader.open/2` and is used to access
  the contents of the file

  It should not be manipulated directly.

  """

  @enforce_keys [:zip_handle, :workbook]
  defstruct zip_handle: nil, workbook: nil

  @typedoc """
  XLSX package
  """
  @type t :: %__MODULE__{
          zip_handle: XlsxReader.ZipArchive.zip_handle(),
          workbook: XlsxReader.Workbook.t()
        }
end

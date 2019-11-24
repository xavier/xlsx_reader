# XlsxReader

An XLSX reader in Elixir.

Features:

- Accepts XLSX data located on the file system or in memory
- Automatic type conversions (numbers, date & times, booleans)
- Optional support for arbitrary precision [decimal](https://github.com/ericmj/decimal) numbers
- Straightforward architecture: no ETS tables, no race-conditions, no manual resource management

## Installation

If [available in Hex](https://hex.pm/docs/publish), the package can be installed
by adding `xlsx_reader` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:xlsx_reader, "~> 0.1.0"}
  ]
end
```

Documentation can be generated with [ExDoc](https://github.com/elixir-lang/ex_doc)
and published on [HexDocs](https://hexdocs.pm). Once published, the docs can
be found at [https://hexdocs.pm/xlsx_reader](https://hexdocs.pm/xlsx_reader).

## Usage

```elixir

{:ok, package} = XlsxReader.open("test.xlsx")

XlsxReader.sheet_names(package)
# ["Sheet 1", "Sheet 2", "Sheet 3"]

{:ok, rows} = XlsxReader.sheet(package, "Sheet 1")
# [
#   ["Date", "Temperature"], 
#   [~D[2019-11-01], 8.4], 
#   [~D[2019-11-02], 7.5], 
#   ...
# ]

{:ok, rows} = XlsxReader.sheet(package, "Sheet 1", number_type: Decimal)
# [
#   ["Date", "Temperature"], 
#   [~D[2019-11-01], %Decimal{coef: 84, exp: -1, sign: 1}], 
#   [~D[2019-11-02], %Decimal{coef: 75, exp: -1, sign: 1}], 
#   ...
# ]

{:ok, sheets} = XlsxReader.sheets(package)
# [
#   {"Sheet 1", [["Date", "Temperature"], ...]}, 
#   {"Sheet 2", [...]}, 
#   ...
# ]
```

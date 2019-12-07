![XlsxReader logo](https://raw.github.com/xavier/xlsx_reader/master/assets/logo.png)

# XlsxReader

An XLSX reader in Elixir.

Features:

- Accepts XLSX data located on the file system or in memory
- Automatic type conversions (numbers, date & times, booleans)
- Optional support for arbitrary precision [decimal](https://github.com/ericmj/decimal) numbers
- Straightforward architecture: no ETS tables, no race-conditions, no manual resource management

The docs can be found at [https://hexdocs.pm/xlsx_reader](https://hexdocs.pm/xlsx_reader).

## Installation

Add `xlsx_reader` as a dependency in your  `mix.exs`:

```elixir
def deps do
  [
    {:xlsx_reader, "~> 0.1.0"}
  ]
end
```

Run `mix deps.get` in your shell to fetch and compile XlsxReader. 

## Examples

### Loading from the file system

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
```

### Loading from memory

```elixir
blob = File.read!("test.xlsx")

{:ok, package} = XlsxReader.open(blob, source: :binary)
```

### Loading all sheets at once

```elixir
{:ok, sheets} = XlsxReader.sheets(package)
# [
#   {"Sheet 1", [["Date", "Temperature"], ...]}, 
#   {"Sheet 2", [...]}, 
#   ...
# ]
```

### Using arbitrary precision numbers

```elixir
{:ok, rows} = XlsxReader.sheet(package, "Sheet 1", number_type: Decimal)
# [
#   ["Date", "Temperature"], 
#   [~D[2019-11-01], %Decimal{coef: 84, exp: -1, sign: 1}], 
#   [~D[2019-11-02], %Decimal{coef: 75, exp: -1, sign: 1}], 
#   ...
# ]
```

## License

Copyright 2019 Xavier Defrang

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

   http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

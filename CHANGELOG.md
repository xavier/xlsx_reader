# Changelog

## [0.8.7] - 2024-07-14

- Fix issue with some empty cell elements being returned as `:expect_formula`, they are now returned as empty strings

## [0.8.6] - 2024-06-28

- Fix handling of cell type/style when some cell elements were missing attributes
- Handle date and datetime values encoded as numeric cell types

## [0.8.5] - 2024-06-02

- Handle files without shared strings or styles relationships.

## [0.8.4] - 2024-04-28

- Upgrade ex_doc
- Fix issue with some empty cell elements being returned as `:expect_chars`, they are now returned as empty strings.

## [0.8.3] - 2024-04-19

- Improve handling of UTF-8/16/32 encoding

## [0.8.2] - 2024-04-06

- Add `exclude_hidden_sheets?` option
- Return `#ERROR` value instead of crashing in case of cell conversion
  error

## [0.8.1] - 2024-01-11

- Add support for shared formulas

## [0.8.0] - 2023-12-11

- Add `cell_data_format` option to return data as `Cell` structs instead of values

## [0.7.0] - 2023-10-15

- Improve ZIP file error handling
- Update Saxy XML parser
- Improve UTF-16 support

## [0.6.0] - 2022-10-30

- Update Saxy XML parser

## [0.5.0] - 2022-06-12

- Require Elixir 1.10 to fix publishing of documentation

## [0.4.3] - 2021-02-08

- Improve compatibility with XLSX writers (Excel for Mac, …) which completely omit empty rows in worksheets

## [0.4.2] - 2021-02-09

- Add `skip_row?` callback

## [0.4.1] - 2020-10-15

- Add support for `decimal ~> 2.0`

## [0.4.0] - 2020-06-23

- Add `:supported_custom_format` option to `XlsxReader.open/2`
- Support ISO 8601 and US date/time custom format by default

## [0.3.0] - 2020-05-07

- Add `:only` and `:except` options to `XlsxReader.sheets/2` and `XlsxReader.async_sheets/3`

## [0.2.0] - 2020-04-27

- Add `XlsxReader.async_sheets/3`

## [0.1.4] - 2020-04-24

- Speed-up shared string and styles lookups

## [0.1.3] - 2020-02-26

- Improve compatibility with XLSX writers (Excel, Elixslx, …) which completely omit empty cells in worksheets

## [0.1.2] - 2019-12-30

- Add `String` number type to disable numeric conversions

## [0.1.1] - 2019-12-20

- Improve handling of whitespace in shared strings

## [0.1.0] - 2019-12-16

- Initial release

# Changelog

## [0.4.0] - unreleased

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

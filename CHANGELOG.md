## 2.0.0 (2026-11-19)

Loiste 2.0 updates to POI 6. Loiste's own API is unchanged, but the new POI version
requires Java 17. Stay on Loiste 1.2.x if you need Java 8 or 11.

- **Breaking:** Requires Java 17 or later (POI 6 is compiled for Java 17)
- **Breaking:** Updated org.apache.poi/poi-ooxml to version 6.0.0
    - Code calling POI directly, e.g. in `:interop` functions, may need changes.
      POI 6 removes deprecated methods, e.g. `Font.getIndexAsInt` (use `getIndex`, which now returns `int`),
      and changes some behavior, e.g. `DataFormatter` uses cached values for formula cells by default.
      See the [POI changelog](https://poi.apache.org/changes.html)

**[compare](https://github.com/metosin/loiste/compare/1.2.0...2.0.0)**

## 1.2.0 (2026-11-19)

- Support writing `java.time` values ([#14](https://github.com/metosin/loiste/issues/14))
    - `LocalDate` and `LocalDateTime` are written as is
    - `Instant` and `ZonedDateTime` are written like `java.util.Date`, as wall-clock time in the JVM default time zone
    - Like `java.util.Date`, use a style with a date `:data-format` to show the cell as a date
- Support all `BorderStyle` and `FillPatternType` values in cell styles ([#9](https://github.com/metosin/loiste/issues/9))
    - Every enum value is available as a kebab-case keyword, e.g. `:medium-dashed`, `:double`, `:fine-dots`
      (`:solid` still means `SOLID_FOREGROUND`)
    - Enum values can also be passed directly, e.g. `:border-bottom BorderStyle/DOTTED`
    - Invalid border values now throw `IllegalArgumentException`, like invalid fill patterns already did
- Allow direct POI calls for things Loiste doesn't support ([#11](https://github.com/metosin/loiste/issues/11))
    - Cell and row maps take an `:interop` function, called with the `Cell` or `Row` and the options map,
      after the values have been written. The options map contains `:workbook`, `:sheet` and the created `:styles`
    - Style and font options take an `:interop` function, called with the `CellStyle` or `Font`
    - Style `:font` accepts a `Font` instance and `write-rows!` `:styles` values accept `CellStyle` instances,
      so fonts and styles can be created once and shared, e.g. between styles and rich text written in `:interop`.
      Instances created in another workbook throw `IllegalArgumentException`
    - Streaming workbooks write strings inline by default, which drops rich text formatting.
      Use `(l/streaming-workbook (l/workbook) {:shared-strings-table? true})` to write rich text.
- Fix writing `org.joda.time.DateTime` values with `loiste.joda-time`:
  the `CellWrite` implementation had the wrong arity and threw `ArityException`.
- Updated org.apache.poi/poi-ooxml to version 5.5.1

**[compare](https://github.com/metosin/loiste/compare/1.1.5...1.2.0)**

## 1.1.5 (2025-01-15)

- Updated org.apache.poi/poi-ooxml to version 5.4.0

**[compare](https://github.com/metosin/loiste/compare/1.1.4...1.1.5)**

## 1.1.4 (2023-12-07)

- Updated org.apache.poi/poi-ooxml to version 5.2.5
- Write `nil` values as BLANK cells, instead of empty string.

**[compare](https://github.com/metosin/loiste/compare/1.1.3...1.1.4)**

## 1.1.3 (2022-12-19)

- Updated org.apache.poi/poi-ooxml to version 5.2.3
    (fixes a CVE in POI)

**[compare](https://github.com/metosin/loiste/compare/1.1.2...1.1.3)**

## 1.1.2 (2022-01-25)

- Updated org.apache.poi/poi-ooxml to version 5.2.0
    (fixes a CVE in POI dependencies)

**[compare](https://github.com/metosin/loiste/compare/1.1.1...1.1.2)**

## 1.1.1 (2021-11-22)

- Updated org.apache.poi/poi-ooxml to version 5.1.0
    (fixes several CVEs in POI dependencies)
    - POI now uses log4j2 for logging, you need to provide a implementation for that.
    - Use org.apache.logging.log4j/log4j-to-slf4j to redirect logging to slf4j.
- Switch `color` function to use rgb byte-array instead of going through `java.awt.Color`
(should not break existing uses)

**[compare](https://github.com/metosin/loiste/compare/1.1.0...1.1.1)**

## 1.1.0 (2021-09-22)

- Updated org.apache.poi/poi-ooxml to version 5.0.0

**[compare](https://github.com/metosin/loiste/compare/1.0.0...1.1.0)**

## 1.0.0 (2020-02-27)

- Update POI dependency and fix compatibility with the latest POI

**[compare](https://github.com/metosin/loiste/compare/0.1.0...1.0.0)**

## 0.1.0 (2017-03-29)

- Uses `isCellDateFormatted` to check if numeric cell value should be
returned as `Date`
- Added `read-sheet` function which uses column headers from the first row
for data map keys
- Supports `TRUE()` and `FALSE()` formulas (coerced to booleans)
- Ignore columns with nil in read spec ([#5](https://github.com/metosin/loiste/issues/5))
- Close input-stream in `-to-workbook` `URL` implementation ([#4](https://github.com/metosin/loiste/issues/4))
- Don't skip empty values at the end of row

**[compare](https://github.com/metosin/loiste/compare/0.0.8...0.1.0)**

## 0.0.8 (2017-01-02)

- Fix cell-style for streaming workbooks (wrong number of arguments given to instance?)

**[compare](https://github.com/metosin/loiste/compare/0.0.7...0.0.8)**

## 0.0.7 (2016-06-01)

- Bugfix for setting top, left and right borders incorrectly to bottom

**[compare](https://github.com/metosin/loiste/compare/0.0.6...0.0.7)**

## 0.0.6 (2016-06-01)

- Support for setting row height
- Support for setting cell wrap option

**[compare](https://github.com/metosin/loiste/compare/0.0.5...0.0.6)**

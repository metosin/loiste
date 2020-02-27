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

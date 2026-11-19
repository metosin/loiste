# Loiste

An Excel library for Clojure, built on [Apache POI](https://poi.apache.org/).

> Verb
> **Loistaa (Excel)**
> 
> 1. (intransitive) To shine, glow, glare, sparkle (to emit light).
> 2. (intransitive) To excel, shine (to distinguish oneself).
> 
> Nouns: Loiste

## Latest version

[![Clojars Project](https://clojars.org/metosin/loiste/latest-version.svg)](https://clojars.org/metosin/loiste)

## Requirements

- Java 17 or later. Loiste 2.x uses POI 6, which requires Java 17.
  Use Loiste 1.2.x for Java 8 and 11.
- A [Log4j 2](https://logging.apache.org/log4j/2.x/) implementation, as POI logs using the Log4j 2 API.
  For example, to send POI logs to SLF4J, add `org.apache.logging.log4j/log4j-to-slf4j`.

## Features

- Create new workbooks, or open existing ones, e.g. to use as a template
- Streaming workbooks (`SXSSFWorkbook`) for writing large files
- Read sheets as maps, with keys from the header row (`read-sheet`)
- Write rows from any seq, including lazy seqs (`write-rows!`)
- Cell styles: fonts, colors, borders, fill patterns and data formats
- Column widths and row heights
- Write `java.util.Date` and `java.time` values, and Joda-Time values with `loiste.joda-time`
- `:interop` functions for direct POI calls, for anything Loiste doesn't support

## Usage

```clj
(require '[loiste.core :as l])
```

### Reading

`read-sheet` uses the first row as column names, and returns the other rows as maps.
Blank rows are left out.

```clj
(let [wb (l/workbook (java.io.File. "people.xlsx"))]
  (l/read-sheet (l/sheet wb 0)))
;; => ({:Name "Alice", :Age 30.0} {:Name "Bob", :Age 25.0})

;; Use :column-name-fn to change how the column names are turned into keys
(l/read-sheet (l/sheet wb "People") {:column-name-fn identity})
;; => ({"Name" "Alice", "Age" 30.0} ...)
```

`l/workbook` takes a `File`, `InputStream` or `URL`. `l/sheet` takes a sheet name or index.
Numbers are read as doubles, and cells with a date format as `java.util.Date`.
Formula cells are read as `nil`, except for `TRUE()` and `FALSE()`.

### Writing

```clj
(let [wb (l/workbook)
      sheet (l/sheet wb "People")]
  (l/write-rows! wb sheet
                 {:styles {:header {:font {:bold true}
                                    :border-bottom :thin}
                           :date {:data-format {:type :date
                                                :locale "fi"
                                                :pattern "d.M.yyyy"}}}
                  :columns [{:width (l/column-width 20)}
                            {:width (l/column-width 12)}]}
                 [[{:value "Name" :style :header} {:value "Born" :style :header}]
                  ["Alice" {:value (java.time.LocalDate/of 1995 3 1) :style :date}]
                  ["Bob" {:value (java.time.LocalDate/of 2000 7 15) :style :date}]])
  (l/to-file! "people.xlsx" wb))
```

Rows are vectors of cell values, or maps with `:values` and an optional `:height` in points.
Cell values are strings, numbers, booleans, keywords, dates or `nil`, or maps
with `:value` and an optional `:style`.

To write a new file from a template, open the template with `l/workbook`, write
to its sheets, and save it to a new file.

### Styles

Style options:

- `:font`: `{:bold true, :name "Arial", :color (.getIndex IndexedColors/RED)}`, or a `Font` created with `l/font`
- `:border-top`, `:border-right`, `:border-bottom`, `:border-left`: a `BorderStyle` keyword, e.g. `:thin` or `:medium-dashed`
- `:fill-pattern`: a `FillPatternType` keyword, e.g. `:solid` or `:fine-dots`
- `:foreground-color`, `:background-color`: `(l/color r g b)`
- `:data-format`: `{:type :date, :locale "fi", :pattern "d.M.yyyy"}`
  or `{:type :custom, :format-str "0.00"}`
- `:wrap`: wrap text

Border and fill pattern keywords are the kebab-case names of the POI enum values,
and the enum values can also be used directly. Dates need a date `:data-format`
to show as dates in Excel.

### Streaming

Streaming workbooks keep only a window of rows in memory, so they can write large files.

```clj
(with-open [wb (l/streaming-workbook)]
  (let [sheet (l/sheet wb "Data")]
    (l/write-rows! wb sheet {} (map (fn [i] [i (str "row " i)]) (range 100000)))
    (l/to-file! "data.xlsx" wb)))
```

Streaming workbooks write the rows to temporary files. Closing the workbook,
e.g. with `with-open`, removes them.

### Direct POI calls

For anything Loiste doesn't support, cell and row maps take an `:interop` function.
It is called with the POI `Cell` or `Row` after the values are written, and with the
options map, which has the `:workbook`, `:sheet` and created `:styles`.
Style and font options also take an `:interop` function, called with the `CellStyle` or `Font`.

```clj
(l/write-rows! wb sheet {}
               [[1 2 {:interop (fn [^org.apache.poi.ss.usermodel.Cell cell _]
                                 (.setCellFormula cell "A1+B1"))}]])
```

Create fonts and styles once, outside `:interop` functions, as they are stored in the workbook.

## License

Copyright © 2015-2026 [Metosin Oy](https://www.metosin.fi)

Distributed under the Eclipse Public License, the same as Clojure.

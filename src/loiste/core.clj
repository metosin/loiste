(ns loiste.core
  (:require [clojure.java.io :as io]
            [clojure.string :as s])
  (:import [java.util Locale]
           [java.awt Color]
           [org.apache.poi.ss.usermodel Workbook Sheet Row Cell CellStyle FillPatternType]
           [org.apache.poi.ss.util DateFormatConverter]
           [org.apache.poi.xssf.usermodel XSSFWorkbook XSSFDataFormat XSSFFont
            XSSFColor]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]))

(set! *warn-on-reflection* true)

(defprotocol ToWorkbook
  (-to-workbook [this] "Reads object to Workbook"))

(extend-protocol ToWorkbook
  Workbook
  (-to-workbook [this]
    this)
  java.io.File
  (-to-workbook [this]
    (doto (XSSFWorkbook. this)
      (.setMissingCellPolicy Row/RETURN_BLANK_AS_NULL)))
  java.io.InputStream
  (-to-workbook [this]
    (doto (XSSFWorkbook. this)
      (.setMissingCellPolicy Row/RETURN_BLANK_AS_NULL)))
  java.net.URL
  (-to-workbook [this]
    (doto (XSSFWorkbook. (io/input-stream this))
      (.setMissingCellPolicy Row/RETURN_BLANK_AS_NULL))))

(defn workbook
  "Creates a new workbook or opens existing workbook.

  Input can be File, InputStream, URL or Workbook."
  ([]
   (XSSFWorkbook.))
  ([input]
   (-to-workbook input)))

(defn streaming-workbook
  "Creates a new streaming workbook or opens existing workbook
  in streaming mode.

  Input can be File, InputStream, URL or Workbook."
  ([]
   (SXSSFWorkbook.))
  ([input]
   (streaming-workbook input nil))
  ([input
    {:keys [row-access-window compress-tmp-files? shared-strings-table?]
     :or {row-access-window 100
          compress-tmp-files? false
          shared-strings-table? false}}]
   (SXSSFWorkbook. (-to-workbook input) row-access-window compress-tmp-files? shared-strings-table?)))

(defn sheet
  "Returns sheet with given name or index, or creates new sheet with
  given name if one doesn't exist."
  ([^Workbook wb]
   (.createSheet wb))
  ([^Workbook wb sheet-name-or-index]
   (or (cond
         (string? sheet-name-or-index) (.getSheet wb ^String sheet-name-or-index)
         (number? sheet-name-or-index) (.getSheetAt wb (int sheet-name-or-index)))
       (if (string? sheet-name-or-index)
         (.createSheet wb ^String sheet-name-or-index)
         (.createSheet wb)))))

;;
;; Reading
;;

(defn rows [^Sheet sheet]
  (seq sheet))

(defn cells [^Row row]
  (if row
    (for [col (range (.getLastCellNum row))]
      (.getCell row col))))

(defn cell-value [^Cell cell]
  (if cell
    ; FIXME: Why doesn't case work here?
    (condp = (.getCellType cell)
      Cell/CELL_TYPE_NUMERIC  (.getNumericCellValue cell)
      Cell/CELL_TYPE_STRING   (.getStringCellValue cell)
      ;Cell/CELL_TYPE_FORMULA  nil
      Cell/CELL_TYPE_BLANK    nil
      Cell/CELL_TYPE_BOOLEAN  (.getBooleanCellValue cell)
      ;Cell/CELL_TYPE_ERROR    nil
      nil)))

(defn- parse-row-spec [data [value [cell-name & tfs]]]
  (assoc data cell-name (if tfs (reduce (fn [v f] (f v)) value tfs) value)))

(defn parse-row
  ([specs ^Row row] (parse-row specs identity row))
  ([specs transform ^Row row]
    (reduce parse-row-spec {} (map vector (->> row cells (map (comp transform cell-value))) specs))))

(defn blank-cells? [row]
  (every? (fn [cell] (and (string? cell) (s/blank? cell))) (vals row)))

(defn strip-blank-rows [rows]
  (filter (complement blank-cells?) rows))

; (defn as-longs [x] (if (instance? Double x) (long x) x))

(defn sheet->map [spec sheet]
  (->> sheet
       (rows)
       (map (partial parse-row spec))
       (strip-blank-rows)))

;;
;; Write to excel
;;

(defprotocol CellWrite
  (-write [value cell options]))

(extend-protocol CellWrite
  java.lang.String
  (-write [value ^Cell cell options]
    (.setCellValue cell value))
  java.lang.Double
  (-write [value ^Cell cell options]
    (.setCellValue cell value))
  java.lang.Boolean
  (-write [value ^Cell cell options]
    (.setCellValue cell value))
  java.util.Date
  (-write [value ^Cell cell options]
    (.setCellValue cell value))
  java.lang.Long
  (-write [value ^Cell cell options]
    (.setCellValue cell (double value)))
  clojure.lang.IPersistentMap
  (-write [value ^Cell cell options]
    (if (:style value)
      (.setCellStyle cell (get-in options [:styles (:style value)])))
    (-write (:value value) cell options))
  nil
  (-write [value ^Cell cell options]
    (.setCellValue cell "")))

(defn write-cell! [^Cell cell {:keys [style] :as options} value]
  (if style
    (.setCellStyle cell (get-in options [:styles style])))
  (-write value cell options))

(defn write-cells! [^Row row options data]
  (doall
    (map (fn [col-num data column-options]
           (write-cell! (.createCell row col-num) (merge options column-options) data))
         (range)
         data
         (concat (:columns options) (repeat nil)))))

(defn font [^Workbook wb {:keys [color name bold]}]
  (doto (.createFont wb)
    (cond-> color (.setColor color))
    (cond-> name (.setName name))
    (cond-> bold (.setBold bold))))

(defn color [r g b]
  (XSSFColor. (Color. (int r) (int g) (int b))))

(defn data-format [^Workbook wb {:keys [type] :as data-format}]
  (case type
    :date (let [l (Locale. (:locale data-format))
                date-fmt (DateFormatConverter/convert l (:pattern data-format))]
            (.getFormat (.createDataFormat wb) date-fmt))
    :custom (let [data-fmt (.createDataFormat wb)]
              (.getFormat data-fmt (:format-str data-format)))))

(def border
  {:thick CellStyle/BORDER_THICK
   :thin CellStyle/BORDER_THIN})

(def fill-pattern
  {:solid FillPatternType/SOLID_FOREGROUND})

; XSSFWorkbook or SXSSFWorkbook
; Workbook doesn't work with XSSFColor
(defn cell-style [wb
                  {:keys [background-color foreground-color
                          border-bottom border-left border-right border-top]
                   :as options}]
  (doto (.createCellStyle wb)
    (cond-> (:font options) (.setFont (font wb (:font options))))
    (cond-> (:fill-pattern options) (.setFillPattern (fill-pattern (:fill-pattern options))))
    (cond-> foreground-color (.setFillForegroundColor foreground-color))
    (cond-> background-color (.setFillBackgroundColor background-color))
    (cond-> border-bottom (.setBorderBottom (border border-bottom)))
    (cond-> border-left (.setBorderBottom (border border-left)))
    (cond-> border-right (.setBorderBottom (border border-right)))
    (cond-> border-top (.setBorderBottom (border border-top)))
    (cond-> (:data-format options) (.setDataFormat (data-format wb (:data-format options))))
    ))

(defn write-rows! [^Workbook wb ^Sheet sheet options rows]
  (let [styles (into {} (for [[k v] (:styles options)]
                          [k (cell-style wb v)]))
        options (assoc options :styles styles)]
    (doseq [[i {:keys [width style]}] (map-indexed vector (:columns options))]
      (if width
        (.setColumnWidth sheet i width))
      ; Doesn't work in Excel -> setCellStyle for each cell
      (if style
        (.setDefaultColumnStyle sheet i (get styles style))))
    (doall
      (map-indexed (fn [row-num data]
                     (if data
                       (write-cells! (.createRow sheet row-num) options data)))
                   rows))))

(defn to-file! [file ^Workbook wb]
  (with-open [os (io/output-stream file)]
    (.write wb os)))

(defn column-width
  "Column width is given in 1/256ths of character width."
  ([num-chars]
   (column-width num-chars 1.25))
  ([num-chars padding]
   (int (* 256 num-chars padding))))

(ns loiste.core
  (:require [clojure.java.io :as io]
            [clojure.string :as s])
  (:import [java.util Locale]
           [java.awt Color]
           [org.apache.poi.ss.usermodel Workbook Sheet Row Cell CellStyle FillPatternType DataFormat CellStyle DateUtil
            Row$MissingCellPolicy BorderStyle]
           [org.apache.poi.ss.util DateFormatConverter]
           [org.apache.poi.xssf.usermodel XSSFWorkbook XSSFColor XSSFCellStyle]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]))

(defprotocol ToWorkbook
  (-to-workbook [this] "Reads object to Workbook"))

(extend-protocol ToWorkbook
  Workbook
  (-to-workbook [this]
    this)
  java.io.File
  (-to-workbook [this]
    (doto (XSSFWorkbook. this)
      (.setMissingCellPolicy Row$MissingCellPolicy/RETURN_BLANK_AS_NULL)))
  java.io.InputStream
  (-to-workbook [this]
    ;; XSSFWorkbook contructor buffers the whole IS into memory, so the stream
    ;; can be closed.
    (with-open [is this]
      (doto (XSSFWorkbook. is)
        (.setMissingCellPolicy Row$MissingCellPolicy/RETURN_BLANK_AS_NULL))))
  java.net.URL
  (-to-workbook [this]
    (-to-workbook (io/input-stream this))))

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

(def static-formulas
  {"TRUE()" true
   "FALSE()" false})

(defn cell-value [^Cell cell]
  (if cell
    (case (.name (.getCellTypeEnum cell))
      ;; Note: createCell + setCellStyle + isCellDateFormatted doesn't work
      ;; isCellDateFormatted seems to require that workboot is written to file or something
      "NUMERIC"  (if (DateUtil/isCellDateFormatted cell)
                   (.getDateCellValue cell)
                   (.getNumericCellValue cell))
      "STRING"   (.getStringCellValue cell)
      "FORMULA"  (let [formula (.getCellFormula cell)]
                   (get static-formulas formula))
      "BOOLEAN"  (.getBooleanCellValue cell)
      "BLANK"    nil
      "ERROR"    nil
      "_NONE"    nil
      )))

(defn- parse-row-spec [data [value spec]]
  (if spec
    (let [[cell-name & tfs] spec]
      (assoc data cell-name (if tfs (reduce (fn [v f] (f v)) value tfs) value)))
    data))

(defn parse-row
  ([specs ^Row row] (parse-row specs identity row))
  ([specs transform ^Row row]
   ;; concat + repeat to make sure there is a value for each spec, else map will missing
   ;; properties from the end of the spec if the last values are empty
   (reduce parse-row-spec {} (map vector (concat (->> row cells (map (comp transform cell-value))) (repeat nil)) specs))))

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

(defn read-sheet
  "Create map key names based on the column labels on the first row."
  {:added "0.1.0"}
  ([sheet] (read-sheet sheet nil))
  ([sheet {:keys [column-name-fn]
           :or {column-name-fn keyword}}]
   (let [[headers & rows] (rows sheet)
         spec (map #(vector (column-name-fn (str %))) headers)]
     (sheet->map spec rows))))

;;
;; Write to excel
;;

(defprotocol CellWrite
  (-write [value cell options]))

(extend-protocol CellWrite
  java.lang.String
  (-write [value ^Cell cell options]
    (.setCellValue cell value))
  java.lang.Number
  (-write [value ^Cell cell options]
    (.setCellValue cell (double value)))
  java.lang.Boolean
  (-write [value ^Cell cell options]
    (.setCellValue cell value))
  java.util.Date
  (-write [value ^Cell cell options]
    (.setCellValue cell value))
  clojure.lang.Named
  (-write [value ^Cell cell options]
    (.setCellValue cell (name value)))
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
    (cond-> name (.setFontName name))
    (cond-> bold (.setBold bold))))

(defn color [r g b]
  (XSSFColor. (Color. (int r) (int g) (int b))))

(defn data-format [^Workbook wb {:keys [type] :as data-format}]
  (case type
    :date (let [l (Locale. (:locale data-format))
                date-fmt (DateFormatConverter/convert l ^String (:pattern data-format))]
            (.getFormat (.createDataFormat wb) date-fmt))
    :custom (let [^DataFormat data-fmt (.createDataFormat wb)]
              (.getFormat data-fmt ^String (:format-str data-format)))))

(def border
  {:thick BorderStyle/THICK
   :thin BorderStyle/THIN})

(def fill-pattern
  {:solid FillPatternType/SOLID_FOREGROUND})

(defn cell-style [^Workbook wb
                  {:keys [background-color foreground-color
                          border-bottom border-left border-right border-top
                          wrap]
                   :as options}]
  (let [^CellStyle cell-style (.createCellStyle wb)
        xssf? (or (instance? XSSFWorkbook wb) (instance? SXSSFWorkbook wb))]
    (if (:font options)
      (.setFont cell-style (font wb (:font options))))
    (if (and (:fill-pattern options) xssf?)
      (if-let [^FillPatternType x (fill-pattern (:fill-pattern options))]
        (.setFillPattern ^XSSFCellStyle cell-style x)
        (throw (IllegalArgumentException. (format "Invalid fill-pattern: %s" (:fill-pattern options))))))
    (if (and foreground-color xssf?)
      (.setFillForegroundColor ^XSSFCellStyle cell-style ^XSSFColor foreground-color))
    (if (and background-color xssf?)
      (.setFillBackgroundColor ^XSSFCellStyle cell-style ^XSSFColor background-color))
    (if border-bottom
      (.setBorderBottom cell-style (border border-bottom)))
    (if border-left
      (.setBorderLeft cell-style (border border-left)))
    (if border-right
      (.setBorderRight cell-style (border border-right)))
    (if border-top
      (.setBorderTop cell-style (border border-top)))
    (if (:data-format options)
      (.setDataFormat cell-style (data-format wb (:data-format options))))
    (if wrap
      (.setWrapText cell-style true))
    cell-style))

(defn write-row! [^Sheet sheet options row-num data]
  (if data
    (let [row (.createRow sheet row-num)]
      (if (map? data)
        (do
          (if (:height data)
            (.setHeightInPoints row (:height data)))
          (write-cells! row options (:values data)))
        (write-cells! row options data)))))

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
      (map-indexed #(write-row! sheet options %1 %2) rows))))

(defn to-file! [file ^Workbook wb]
  (with-open [os (io/output-stream file)]
    (.write wb os)))

(defn column-width
  "Column width is given in 1/256ths of character width."
  ([num-chars]
   (column-width num-chars 1.25))
  ([num-chars padding]
   (int (* 256 num-chars padding))))

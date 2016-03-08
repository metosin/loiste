(ns loiste.core
  (:require [clojure.java.io :as io]
            [clojure.string :as s])
  (:import [java.util Locale]
           [org.apache.poi.ss.usermodel Workbook Sheet Row Cell]
           [org.apache.poi.ss.util DateFormatConverter]
           [org.apache.poi.xssf.usermodel XSSFWorkbook XSSFDataFormat]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]))

(set! *warn-on-reflection* true)

(defn blank-cells? [row]
  (every? (fn [cell] (and (string? cell) (s/blank? cell))) (vals row)))

(defn strip-blank-rows [rows]
  (filter (complement blank-cells?) rows))

(defn load-excel [^java.io.InputStream in]
    (doto (XSSFWorkbook. in)
      (.setMissingCellPolicy Row/RETURN_BLANK_AS_NULL)))

(defn load-excel-resource [resource]
  (->> resource
       (io/resource)
       (io/input-stream)
       (load-excel)))

(defn streaming-workbook
  ([^Workbook template] (streaming-workbook template nil))
  ([^Workbook template
    {:keys [row-access-window compress-tmp-files? shared-strings-table?]
     :or {row-access-window 100
          compress-tmp-files? false
          shared-strings-table? false}}]
   (SXSSFWorkbook. template row-access-window compress-tmp-files? shared-strings-table?)))

(defn sheet ^org.apache.poi.ss.usermodel.Sheet [sheet ^Workbook wb]
  (cond
    (string? sheet) (.getSheet wb ^String sheet)
    (number? sheet) (.getSheetAt wb (int sheet))))

(defn rows [^Sheet sheet]
  (seq sheet))

(defn cells [^Row row]
  (if row
    (for [col (range (.getLastCellNum row))]
      (.getCell row col))))

(defn cell-value [^Cell cell]
  (if cell
    (condp = (.getCellType cell)
      Cell/CELL_TYPE_STRING   (.getStringCellValue cell)
      Cell/CELL_TYPE_NUMERIC  (.getNumericCellValue cell)
      Cell/CELL_TYPE_BOOLEAN  (.getBooleanCellValue cell)
      Cell/CELL_TYPE_BLANK    nil)))

(defn- parse-row-spec [data [value [cell-name & tfs]]]
  (assoc data cell-name (if tfs (reduce (fn [v f] (f v)) value tfs) value)))

(defn parse-row
  ([specs ^Row row] (parse-row specs identity row))
  ([specs transform ^Row row]
    (reduce parse-row-spec {} (map vector (->> row cells (map (comp transform cell-value))) specs))))

(defn as-longs [x] (if (instance? Double x) (long x) x))

(defn sheet->map [spec sheet]
  (->> sheet
       (rows)
       (map (partial parse-row spec))
       (strip-blank-rows)))

;;
;; Write to excel
;;

(def fi-locale (Locale. "fi"))
(def date-format (DateFormatConverter/convert ^Locale fi-locale "dd.MM.yyyy HH:mm"))

(defn create-styles! [^Workbook wb]
  {:date-style (doto (.createCellStyle wb)
                 (.setDataFormat (.getFormat (.createDataFormat wb) ^String date-format)))})

(defprotocol CellWrite
  (-write [value cell]))

(extend-protocol CellWrite
  java.lang.String
  (-write [value ^Cell cell]
    (.setCellValue cell value))
  java.lang.Double
  (-write [value ^Cell cell]
    (.setCellValue cell value))
  java.lang.Boolean
  (-write [value ^Cell cell]
    (.setCellValue cell value))
  java.util.Date
  (-write [value ^Cell cell]
    (.setCellValue cell value))
  java.lang.Long
  (-write [value ^Cell cell]
    (.setCellValue cell (double value)))
  nil
  (-write [value ^Cell cell]
    (.setCellValue cell "")))

(defn write-cell! [^Cell cell value]
  (-write value cell))

(defn write-cells! [^Row row data columns styles]
  (doall
    (map (fn [col-num data column]
           (let [cell (.createCell row col-num)]
             (if (:style column)
               (.setCellStyle cell (get styles (:style column))))
             (write-cell! cell data)))
         (range)
         data
         (concat columns (repeat nil)))))

(defn write-rows! [^Sheet sheet rows columns styles]
  (doall
    (map-indexed (fn [row-num data]
                   (-> (.createRow sheet (inc row-num))
                       (write-cells! data columns styles)))
                 rows)))

(defn append-row! [^Sheet sheet data]
  (let [last-row-num (inc (.getLastRowNum sheet))]
    (-> (.createRow sheet last-row-num)
        (write-cells! data))))

(defn autosize-cols!
  "Bloody slow! Do not use ever!"
  [^Sheet sheet]
  (let [num-cols (.getLastCellNum (.getRow sheet 0))]
    (doseq [n (range num-cols)]
      (.autoSizeColumn sheet n))))

(defn column-width
  "Column width is given in 1/256ths of character width."
  ([num-chars]
   (column-width num-chars 1.25))
  ([num-chars padding]
   (int (* 256 num-chars padding))))

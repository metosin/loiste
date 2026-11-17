(ns loiste.core-test
  (:require [clojure.test :refer [deftest testing is]]
            [loiste.core :as l]
            [clojure.java.io :as io])
  (:import [java.io File]
           [java.time Instant LocalDate LocalDateTime ZonedDateTime ZoneId]
           [java.util Date TimeZone]
           [org.apache.poi.ss.usermodel BorderStyle Cell CellStyle FillPatternType Font HorizontalAlignment IndexedColors Row]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]
           [org.apache.poi.xssf.usermodel XSSFRichTextString]))

(def test-spec
  [[:A]
   [:B]
   nil
   [:C]])

(def test-file (io/resource "excel_test.xlsx"))

(def sheet1-result
  [{:A "A2" :B nil  :C 1.0}
   {:A "A3" :B nil  :C 2.0}
   {:A "A4" :B false :C 3.5}
   {:A "A5" :B true  :C 3.0}])

(deftest load-excel-file-test
  (with-open [wb (l/workbook test-file)]
    (testing "there should be correct amount of sheets in the test file"
      (is (= 4 (.getNumberOfSheets wb))))

    (testing "sheet should open the correct sheet"
      (is (= 4 (.getLastRowNum (l/sheet wb "Sheet1")))))

    (testing "parsing excel should return sequence of maps containing data keyed by spec"
      (is (= sheet1-result (rest (l/sheet->map test-spec (l/sheet wb "Sheet1"))))))))

(deftest write-to-template-test
  (with-open [wb (l/workbook (io/resource "excel_test.xlsx"))]
    (let [sheet (l/sheet wb "EmptySheet")
          file (File/createTempFile "test-file" ".xlsx")]
      (l/write-rows!
        wb
        sheet
        {}
        [nil ; Skip header row
         ["Foo" true "ignore" 1.0]
         ["Bar" nil  "ignore" 2.0]])

      (l/to-file! file wb)

      (testing "Writing results into a correct number of rows"
        (is (= 2 (.getLastRowNum sheet))))

      (testing "parse the written data"
        (is (= [{:A "Foo" :B true :C 1.0}
                {:A "Bar" :B nil   :C 2.0}]
               (rest (l/sheet->map test-spec sheet)))))

      file)))

(deftest strip-blank-rows-test
  (is (= [{:a 1 :b 2 :c 3}
          {:a 2 :b 3 :c 4}]
         (l/strip-blank-rows [{:a 1 :b 2 :c 3}
                              {:a "" :b "" :c ""}
                              {:a 2 :b 3 :c 4}]))))

(deftest write-new-file-test
  ;; Keep the file around after tests for testing opening,
  ;; but remove on the start.
  (let [old-tz (TimeZone/getDefault)
        file (io/file "test-out" "new-file-test.xslx")]
    (try
      (TimeZone/setDefault (TimeZone/getTimeZone "Europe/Helsinki"))
      (io/make-parents file)
      (doseq [wb* [(l/workbook)
                   (l/streaming-workbook)]]
        (.delete file)
        (with-open [wb wb*]
          (let [sheet (l/sheet wb)]
            (println "Test file " (.getPath file))
            (l/write-rows!
              wb
              sheet
              {:styles {:date {:data-format {:type :date :pattern "dd.MM.yyyy HH.MM" :locale "fi"}}
                        :currency {:data-format {:type :custom :format-str "#0\\,00 €"}}
                        :header {:font {:bold true}
                                 :border-bottom :thick
                                 :border-right :thin
                                 :fill-pattern :solid
                                 :foreground-color (l/color 200 200 200)}
                        :wrapping-text {:wrap true}}
               :columns [{:width (l/column-width 10)}
                         {:width (l/column-width 15) :style :date}
                         {:width (l/column-width 10) :style :currency}]}
              [[{:style :header :value "Stuff"}
                {:style :header :value "Date"}
                {:style :header :value "Money"}]
               {:height 100
                :values [{:style :wrapping-text
                          :value "This is a long text that should wrap"}
                         #inst "2016-03-09T14:05:00"]}
               ["Bar" #inst "2016-03-09T14:05:00" 15]
               ["Qux" 100000 100]
               ["Integer" (count '(1 2 3)) 10000]
               ["Keyword" :foo nil]])

            (l/to-file! file wb))))

      (with-open [wb (l/workbook file)]
        (let [read-sheet (l/sheet wb 0)]
          (is (= [{:Stuff "This is a long text that should wrap" :Date #inst "2016-03-09T14:05:00" :Money nil}
                  {:Stuff "Bar" :Date #inst "2016-03-09T14:05:00" :Money 15.0}
                  {:Stuff "Qux" :Date #inst "2173-10-13T21:00:00.000-00:00" :Money 100.0}
                  {:Stuff "Integer" :Date #inst "1900-01-02T22:20:11.000-00:00" :Money 10000.0}
                  {:Stuff "Keyword" :Date "foo" :Money nil}]
                 (l/read-sheet read-sheet)))))
      (finally
        (TimeZone/setDefault old-tz)))))

;; See CellWrite: java.time values without a zone are written as is,
;; Instant and ZonedDateTime are written like java.util.Date, i.e. as
;; wall-clock time in the JVM default time zone.
(deftest write-java-time-test
  (let [old-tz (TimeZone/getDefault)]
    (try
      (TimeZone/setDefault (TimeZone/getTimeZone "Europe/Helsinki"))
      (with-open [wb (l/workbook)]
        (let [sheet (l/sheet wb)
              instant (Instant/parse "2016-03-09T14:05:00Z")]
          (l/write-rows! wb sheet
                         {:styles {:date {:data-format {:type :date :pattern "dd.MM.yyyy HH.mm" :locale "en"}}}}
                         [["LocalDate" "LocalDateTime" "Instant" "ZonedDateTime"]
                          [{:style :date :value (LocalDate/of 2016 3 9)}
                           {:style :date :value (LocalDateTime/of 2016 3 9 14 5)}
                           {:style :date :value instant}
                           {:style :date :value (ZonedDateTime/of 2016 3 9 14 5 0 0 (ZoneId/of "UTC"))}]])
          (testing "values are written as date cells"
            (is (= [(LocalDateTime/of 2016 3 9 0 0)
                    (LocalDateTime/of 2016 3 9 14 5)
                    (LocalDateTime/of 2016 3 9 16 5)
                    (LocalDateTime/of 2016 3 9 16 5)]
                   (map #(.getLocalDateTimeCellValue %) (l/cells (.getRow sheet 1))))))
          (testing "Instant and ZonedDateTime read back as the same point in time"
            (is (= {:Instant (Date/from instant)
                    :ZonedDateTime (Date/from instant)}
                   (select-keys (first (l/read-sheet sheet)) [:Instant :ZonedDateTime]))))))
      (finally
        (TimeZone/setDefault old-tz)))))

(deftest color-test
  (is (= (unchecked-byte 200)
         (aget (.getRGB (l/color 200 200 200)) 0))))

(deftest cell-style-border-and-fill-test
  (with-open [wb (l/workbook)]
    (testing "existing keywords keep working"
      (let [style (l/cell-style wb {:border-bottom :thick
                                    :border-top :thin
                                    :fill-pattern :solid})]
        (is (= BorderStyle/THICK (.getBorderBottom style)))
        (is (= BorderStyle/THIN (.getBorderTop style)))
        (is (= FillPatternType/SOLID_FOREGROUND (.getFillPattern style)))))

    (testing "all enum values are available as kebab-case keywords"
      (let [style (l/cell-style wb {:border-bottom :medium-dashed
                                    :border-left :double
                                    :border-right :dash-dot-dot
                                    :border-top :hair
                                    :fill-pattern :fine-dots})]
        (is (= BorderStyle/MEDIUM_DASHED (.getBorderBottom style)))
        (is (= BorderStyle/DOUBLE (.getBorderLeft style)))
        (is (= BorderStyle/DASH_DOT_DOT (.getBorderRight style)))
        (is (= BorderStyle/HAIR (.getBorderTop style)))
        (is (= FillPatternType/FINE_DOTS (.getFillPattern style))))
      (is (= (count (BorderStyle/values)) (count l/border)))
      (is (= (inc (count (FillPatternType/values))) (count l/fill-pattern))))

    (testing "enum values can be given directly"
      (let [style (l/cell-style wb {:border-bottom BorderStyle/DOTTED
                                    :fill-pattern FillPatternType/BRICKS})]
        (is (= BorderStyle/DOTTED (.getBorderBottom style)))
        (is (= FillPatternType/BRICKS (.getFillPattern style)))))

    (testing "invalid values throw"
      (is (thrown-with-msg? IllegalArgumentException #"Invalid border-bottom: :foo"
                            (l/cell-style wb {:border-bottom :foo})))
      (is (thrown-with-msg? IllegalArgumentException #"Invalid fill-pattern: :foo"
                            (l/cell-style wb {:fill-pattern :foo}))))))

(deftest interop-test
  (doseq [wb* [(l/workbook) (l/streaming-workbook)]]
    (with-open [wb wb*]
      (let [sheet (l/sheet wb)
            fonts-before (.getNumberOfFonts wb)
            ;; Workbook level objects are created once and closed over
            red (l/font wb {:color (.getIndex IndexedColors/RED)})
            green (l/font wb {:color (.getIndex IndexedColors/GREEN)})
            header-font (l/font wb {:bold true
                                    :interop (fn [^Font font] (.setItalic font true))})
            right-aligned (doto (.createCellStyle wb)
                            (.setAlignment HorizontalAlignment/RIGHT))
            rich-text (fn [^Cell cell _]
                        (.setCellValue cell (doto (XSSFRichTextString.)
                                              (.append "foo" red)
                                              (.append "bar" green))))]
        (l/write-rows!
          wb sheet
          {:styles {:header {:font header-font}
                    :right right-aligned
                    :wrap {:interop (fn [^CellStyle style] (.setWrapText style true))}}}
          [[{:style :header :value "A"} {:style :right :value "B"} "C"]
           {:interop (fn [^Row row {:keys [workbook sheet styles]}]
                       (is (identical? wb workbook))
                       (is (some? sheet))
                       (is (identical? right-aligned (:right styles)))
                       (.setHeightInPoints row 40))
            :values [{:interop rich-text}
                     {:style :wrap :interop rich-text}
                     {:value 1}]}
           [{:interop rich-text}
            {:interop (fn [^Cell cell {:keys [styles]}]
                        (.setCellStyle cell (:right styles)))
             :value "x"}
            {:interop (fn [^Cell cell _] (.setCellFormula cell "C2+1"))}]])

        (testing "fonts are only created up front, not for each cell"
          (is (= (+ fonts-before 3) (.getNumberOfFonts wb))))

        (testing "style :font accepts Font instance, font :interop is called"
          (let [font (.getFontAt wb (.getFontIndex (.getCellStyle (.getCell (.getRow sheet 0) 0))))]
            (is (.getBold font))
            (is (.getItalic font))))

        (testing ":styles accepts CellStyle instance"
          (is (= HorizontalAlignment/RIGHT (.getAlignment (.getCellStyle (.getCell (.getRow sheet 0) 1))))))

        (testing "style :interop is called"
          (is (.getWrapText (.getCellStyle (.getCell (.getRow sheet 1) 1)))))

        (testing "row :interop is called"
          (is (= 40.0 (double (.getHeightInPoints (.getRow sheet 1))))))

        (testing "cell :interop is called after value is written"
          (is (= "foobar" (.getStringCellValue (.getCell (.getRow sheet 2) 0))))
          (is (= "x" (.getStringCellValue (.getCell (.getRow sheet 2) 1))))
          (is (= HorizontalAlignment/RIGHT (.getAlignment (.getCellStyle (.getCell (.getRow sheet 2) 1)))))
          (is (= "C2+1" (.getCellFormula (.getCell (.getRow sheet 2) 2)))))

        ;; Streaming workbook logs a warning if closed without being written or disposed
        (when (instance? SXSSFWorkbook wb)
          (.dispose ^SXSSFWorkbook wb))))))

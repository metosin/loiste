(ns loiste.core-test
  (:require [clojure.test :refer [deftest testing is]]
            [loiste.core :as l]
            [clojure.java.io :as io])
  (:import [java.io File]))

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
                {:A "Bar" :B ""   :C 2.0}]
               (rest (l/sheet->map test-spec sheet)))))

      file)))

(deftest strip-blank-rows-test
  (is (= [{:a 1 :b 2 :c 3}
          {:a 2 :b 3 :c 4}]
         (l/strip-blank-rows [{:a 1 :b 2 :c 3}
                              {:a "" :b "" :c ""}
                              {:a 2 :b 3 :c 4}]))))

(deftest write-new-file-test
  (let [file (File/createTempFile "test-file" ".xlsx")]
    (with-open [wb (l/workbook)]
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

        (l/to-file! file wb)))

    (with-open [wb (l/workbook file)]
      (let [read-sheet (l/sheet wb 0)]
        (is (= [{:Stuff "This is a long text that should wrap" :Date #inst "2016-03-09T14:05:00" :Money nil}
                {:Stuff "Bar" :Date #inst "2016-03-09T14:05:00" :Money 15.0}
                {:Stuff "Qux" :Date #inst "2173-10-13T21:00:00.000-00:00" :Money 100.0}
                {:Stuff "Integer" :Date #inst "1900-01-02T22:20:11.000-00:00" :Money 10000.0}
                {:Stuff "Keyword" :Date "foo" :Money ""}]
               (l/read-sheet read-sheet)))))))

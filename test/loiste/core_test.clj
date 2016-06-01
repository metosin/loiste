(ns loiste.core-test
  (:require [clojure.test :refer :all]
            [loiste.core :as l]
            [clojure.java.io :as io])
  (:import [java.io File]))

(def test-spec
  [[:A]
   [:B]
   [:C]])

(def test-workbook (l/workbook (io/resource "excel_test.xlsx")))

(def sheet1-result
  [{:A "A2" :B nil  :C 1.0}
   {:A "A3" :B nil  :C 2.0}
   {:A "A4" :B false :C 3.5}
   {:A "A5" :B true  :C 3.0}])

(deftest load-excel-file-test
  (testing "there should be correct amount of sheets in the test file"
    (is (= 4
           (.getNumberOfSheets test-workbook))))

  (testing "sheet should open the correct sheet"
    (is (= 4
           (.getLastRowNum (l/sheet test-workbook "Sheet1")))))

  (testing "parsing excel should return sequence of maps containing data keyed by spec"
    (is (= sheet1-result (rest (l/sheet->map test-spec (l/sheet test-workbook "Sheet1")))))))

(deftest write-to-template-test
  (let [test-workbook (l/workbook (io/resource "excel_test.xlsx"))
        sheet (l/sheet test-workbook "EmptySheet")
        file (File/createTempFile "test-file" ".xlsx")]
    (l/write-rows!
      test-workbook
      sheet
      {}
      [nil ; Skip header row
       ["Foo" true 1.0]
       ["Bar" nil  2.0]])

    (l/to-file! file test-workbook)

    (testing "Writing results into a correct number of rows"
      (is (= 2 (.getLastRowNum sheet))))

    (testing "Writing results into a correct data getting written"
      (is (= [{:A "Foo" :B true :C 1.0}
              {:A "Bar" :B ""   :C 2.0}]
             (rest (l/sheet->map test-spec sheet)))))

    file))

(deftest strip-blank-rows-test
  (is (= [{:a 1 :b 2 :c 3}
          {:a 2 :b 3 :c 4}]
         (l/strip-blank-rows [{:a 1 :b 2 :c 3}
                              {:a "" :b "" :c ""}
                              {:a 2 :b 3 :c 4}]))))

(deftest write-new-file-test
  (let [test-workbook (l/workbook)
        sheet (l/sheet test-workbook)
        file (File/createTempFile "test-file" ".xlsx")]
    (l/write-rows!
      test-workbook
      sheet
      {:styles {:date {:data-format {:type :date :pattern "dd.MM.yyyy HH.MM" :locale "fi"}}
                :currency {:data-format {:type :custom :format-str "#0\\,00 €"}}
                :header {:font {:bold true}
                         :border-bottom :thick
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
       ["Bar" #inst "2016-03-09T14:05:00"]
       ["Qux" 100000]])
    (l/to-file! file test-workbook)
    file))

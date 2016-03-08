(ns loiste.core-test
  (:require [clojure.test :refer :all]
            [loiste.core :refer :all]))

(def test-spec
  [[:A]
   [:B]
   [:C]])

(def test-workbook (load-excel-resource "excel_test.xlsx"))

(def sheet1-result
  [{:A "A2" :B nil  :C 1.0}
   {:A "A3" :B nil  :C 2.0}
   {:A "A4" :B false :C 3.5}
   {:A "A5" :B true  :C 3.0}])

(def test-rows [["Foo" true 1.0]
                ["Bar" nil  2.0]])

(def write-result
  [{:A "Foo" :B true :C 1.0}
   {:A "Bar" :B ""   :C 2.0}])

(deftest load-excel-file-test
  (testing "there should be correct amount of sheets in the test file"
    (is (= 4
           (.getNumberOfSheets test-workbook))))

  (testing "sheet should open the correct sheet"
    (is (= 4
           (.getLastRowNum (sheet "Sheet1" test-workbook)))))

  (testing "parsing excel should return sequence of maps containing data keyed by spec"
    (is (= sheet1-result (rest (sheet->map test-spec (sheet "Sheet1" test-workbook)))))))

(deftest write-test
  (let [test-workbook (load-excel-resource "excel_test.xlsx")
        write-sheet (sheet "EmptySheet" test-workbook)]
    (write-rows! write-sheet test-rows nil nil)

    (testing "Writing results into a correct number of rows"
      (is (= (count test-rows) (.getLastRowNum write-sheet))))

    (testing "Writing results into a correct data getting written"
      (is (= write-result (rest (sheet->map test-spec write-sheet)))))))

(deftest strip-blank-rows-test
  (is (= [{:a 1 :b 2 :c 3}
          {:a 2 :b 3 :c 4}]
         (strip-blank-rows [{:a 1 :b 2 :c 3}
                            {:a "" :b "" :c ""}
                            {:a 2 :b 3 :c 4}]))))


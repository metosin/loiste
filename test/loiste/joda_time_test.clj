(ns loiste.joda-time-test
  (:require [clojure.test :refer [deftest is]]
            [loiste.core :as l]
            [loiste.joda-time])
  (:import [org.joda.time DateTime DateTimeZone]))

(deftest write-joda-date-time-test
  (with-open [wb (l/workbook)]
    (let [sheet (l/sheet wb)
          dt (DateTime. 2016 3 9 14 5 DateTimeZone/UTC)]
      (l/write-rows! wb sheet
                     {:styles {:date {:data-format {:type :date :pattern "dd.MM.yyyy HH.mm" :locale "en"}}}}
                     [["Date"] [{:style :date :value dt}]])
      (is (= [{:Date (.toDate dt)}]
             (l/read-sheet sheet))))))

(ns loiste.joda-time
  (:require [loiste.core :as loiste])
  (:import [org.joda.time DateTime]
           [org.apache.poi.ss.usermodel Cell]))

(extend-protocol loiste/CellWrite
  org.joda.time.DateTime
  (-write [^DateTime value ^Cell cell]
    (.setCellValue cell (.toDate value))))

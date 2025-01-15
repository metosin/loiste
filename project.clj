(defproject metosin/loiste "1.1.5"
  :description "An Excel library for Clojure"
  :url "https://github.com/metosin/loiste"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.apache.poi/poi-ooxml "5.4.0"]]

  :profiles {:dev {:dependencies [[org.clojure/clojure "1.12.0"]
                                  [joda-time "2.13.0"]

                                  ;; POI needs log4j implementation
                                  [org.apache.logging.log4j/log4j-to-slf4j "2.24.3"]
                                  [ch.qos.logback/logback-classic "1.5.16" :exclusions [org.slf4j/slf4j-api]]
                                  [org.slf4j/slf4j-api "2.0.16"]]
                   :resource-paths ["dev-resources"]}})

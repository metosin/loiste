(defproject metosin/loiste "1.1.4"
  :description "An Excel library for Clojure"
  :url "https://github.com/metosin/loiste"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.apache.poi/poi-ooxml "5.2.5"]]

  :profiles {:dev {:dependencies [[org.clojure/clojure "1.11.1"]
                                  [joda-time "2.12.2"]

                                  ;; POI needs log4j implementation
                                  [org.apache.logging.log4j/log4j-to-slf4j "2.19.0"]
                                  [ch.qos.logback/logback-classic "1.4.5" :exclusions [org.slf4j/slf4j-api]]
                                  [org.slf4j/slf4j-api "1.7.36"]]
                   :resource-paths ["dev-resources"]}})

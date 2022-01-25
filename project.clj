(defproject metosin/loiste "1.1.2"
  :description "An Excel library for Clojure"
  :url "https://github.com/metosin/loiste"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.apache.poi/poi-ooxml "5.2.0"]]

  :profiles {:dev {:dependencies [[org.clojure/clojure "1.10.3"]
                                  [joda-time "2.10.13"]
                                  [nvd-clojure/nvd-clojure "1.9.0"]

                                  ;; POI needs log4j implementation
                                  [org.apache.logging.log4j/log4j-to-slf4j "2.17.1"]
                                  [ch.qos.logback/logback-classic "1.2.10" :exclusions [org.slf4j/slf4j-api]]
                                  [org.slf4j/slf4j-api "1.7.34"]]
                   :resource-paths ["dev-resources"]}})

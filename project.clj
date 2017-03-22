(defproject metosin/loiste "0.0.9-SNAPSHOT"
  :description "An Excel library for Clojure"
  :url "https://github.com/metosin/loiste"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.apache.poi/poi-ooxml "3.14"]]

  :profiles {:dev {:dependencies [[org.clojure/clojure "1.8.0"]
                                  [joda-time "2.9.3"]]
                   :resource-paths ["dev-resources"]}})

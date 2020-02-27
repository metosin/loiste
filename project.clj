(defproject metosin/loiste "1.0.0"
  :description "An Excel library for Clojure"
  :url "https://github.com/metosin/loiste"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.apache.poi/poi-ooxml "4.1.2"]]

  :profiles {:dev {:dependencies [[org.clojure/clojure "1.10.1"]
                                  [joda-time "2.10.5"]]
                   :resource-paths ["dev-resources"]}})

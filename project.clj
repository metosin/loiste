(defproject metosin/loiste "1.1.1"
  :description "An Excel library for Clojure"
  :url "https://github.com/metosin/loiste"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.apache.poi/poi-ooxml "5.1.0"]]

  :profiles {:dev {:dependencies [[org.clojure/clojure "1.10.3"]
                                  [joda-time "2.10.13"]
                                  [nvd-clojure/nvd-clojure "1.9.0"]]
                   :resource-paths ["dev-resources"]}})

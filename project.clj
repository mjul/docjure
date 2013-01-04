(defproject dk.ative/docjure "1.6.0-SNAPSHOT"
  :description "Easily read and write Office documents from Clojure."
  :url "http://github.com/ative/docjure"
  :dependencies [[org.clojure/clojure "1.3.0"]
		 [org.apache.poi/poi "3.9"]
		 [org.apache.poi/poi-ooxml "3.9"]]
  :plugins [[lein-difftest "2.0.0"]]
  :profiles {:1.3 {:dependencies [[org.clojure/clojure "1.3.0"]]}
             :1.4 {:dependencies [[org.clojure/clojure "1.4.0"]]}
             :1.5 {:dependencies [[org.clojure/clojure "1.5.0-RC1"]]}}
  :aliases {"all" ["with-profile" "1.3:1.4:1.5"]})

(defproject com.vijaykiran/docjure "1.7.0"
  :description "Easily read and write Office documents from Clojure."
  :url "http://github.com/vijaykiran/docjure"
  :license {:name "MIT License"
            :url "http://http://en.wikipedia.org/wiki/MIT_License"}
  :dependencies [[org.clojure/clojure "1.5.1"]
		 [org.apache.poi/poi "3.9"]
		 [org.apache.poi/poi-ooxml "3.9"]]
  :plugins [[lein-difftest "2.0.0"]]
  :profiles {:1.3 {:dependencies [[org.clojure/clojure "1.3.0"]]}
             :1.4 {:dependencies [[org.clojure/clojure "1.4.0"]]}
             :1.5 {:dependencies [[org.clojure/clojure "1.5.1"]]}}
  :aliases {"all" ["with-profile" "1.3:1.4:1.5"]}
  :global-vars {*warn-on-reflection* true})

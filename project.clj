(defproject dk.ative/docjure "1.5.1-SNAPSHOT"
  :description "Easily read and write Office documents from Clojure."
  :url "http://github.com/ative/docjure"
  :dependencies [[org.clojure/clojure "1.2.0"]
                 [org.clojure/clojure-contrib "1.2.0"]
		 [org.apache.poi/poi "3.8-beta4"]
		 [org.apache.poi/poi-ooxml "3.8-beta4"]]
  :dev-dependencies [[swank-clojure "1.3.0-SNAPSHOT"]
		     [lein-clojars "0.6.0"]
		     [lein-difftest "1.3.2-SNAPSHOT"]]
  :hooks [leiningen.hooks.difftest]
  )

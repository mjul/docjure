(defproject dk.ative/docjure "1.5.0-SNAPSHOT"
  :description "Easily read and write Office documents from Clojure."
  :url "http://github.com/ative/docjure"
  :dependencies [[org.clojure/clojure "1.2.0"]
                 [org.clojure/clojure-contrib "1.2.0"]
		 [org.apache.poi/poi "3.6"]
		 [org.apache.poi/poi-ooxml "3.6"]]
  :dev-dependencies [[swank-clojure "1.3.0-SNAPSHOT"]
		     [lein-clojars "0.5.0"]
		     [lein-difftest "1.3.0"]]
  :hooks [leiningen.hooks.difftest]
  )

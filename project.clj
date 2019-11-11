(defproject dk.ative/docjure "1.15.0-SNAPSHOT"
  :description "Easily read and write Office documents from Clojure."
  :url "http://github.com/mjul/docjure"
  :license {:name "MIT License"
            :url "http://http://en.wikipedia.org/wiki/MIT_License"}
  :dependencies [[org.clojure/clojure "1.10.1"]
                 [org.apache.poi/poi "4.1.1"]
                 [org.apache.poi/poi-ooxml "4.1.1"]]
  :plugins [[lein-difftest "2.0.0"]]
  :profiles { :1.9  { :dependencies [[org.clojure/clojure "1.9.0"]]}
              :1.10 { :dependencies [[org.clojure/clojure "1.10.1"]]}
              :test { :global-vars  {*warn-on-reflection* false}
                      :source-paths ["src" "test"]}}
  :aliases {"all" ["with-profile" "1.9:1.10"]}
  :global-vars {*warn-on-reflection* true})


(defproject dk.ative/docjure "1.22.0"
  :description "Easily read and write Office documents from Clojure."
  :url "http://github.com/mjul/docjure"
  :license {:name "MIT License"
            :url "http://http://en.wikipedia.org/wiki/MIT_License"}
  :dependencies [[org.clojure/clojure "1.12.0"]
                 [org.apache.poi/poi "5.4.1"]
                 [org.apache.poi/poi-ooxml "5.4.1"]]
  :plugins [[lein-difftest "2.0.0"]
            [lein-nvd "1.4.0"]
            [lein-ancient "0.6.15"]
            [com.github.clj-kondo/lein-clj-kondo "0.2.5"]
            [dev.weavejester/lein-cljfmt "0.13.0"]]
  :profiles {:1.5  {:dependencies [[org.clojure/clojure "1.5.1"]]}
             :1.6  {:dependencies [[org.clojure/clojure "1.6.0"]]}
             :1.7  {:dependencies [[org.clojure/clojure "1.7.0"]]}
             :1.8  {:dependencies [[org.clojure/clojure "1.8.0"]]}
             :1.9  {:dependencies [[org.clojure/clojure "1.9.0"]]}
             :1.10 {:dependencies [[org.clojure/clojure "1.10.3"]]}
             :1.11 {:dependencies [[org.clojure/clojure "1.11.3"]]}
             :1.12 {:dependencies [[org.clojure/clojure "1.12.0"]]}
             :test {:global-vars  {*warn-on-reflection* false}
                    :resource-paths ["test/dk/ative/docjure/testdata"]}}
  :aliases {"all" ["with-profile" "1.5:1.6:1.7:1.8:1.9:1.10:1.11:1.12"]}
  :global-vars {*warn-on-reflection* true})


# Docjure

Docjure makes reading and writing Office documents in Clojure easy.
It was written by Martin Jul using the Apache POI library.


## Usage

### Example: Read a Price List sheet

    (use 'dk.ative.docjure.spreadsheet)       

    ; Load a spreadsheet and read the first two columns from the 
    ; price list sheet:
    (->> (load-workbook "spreadsheet.xlsx")
         (select-sheet "Price List")
         (select-columns {:A :name, :B :price}))

    > [{:name "Foo Widget", :price 100}, {:name "Bar Widget", :price 200}]

### Automatically get the Docjure jar from Clojars

The Docjure jar is distributed on [Clojars](http://clojars.org/dk.ative/docjure). 

If you are using the Leiningen build tool just add this line to the
:dependencies list in project.clj to use it:

    [dk.ative/docjure "1.3.0"]	

Remember to issue the 'lein deps' command to download it.

#### Example project.clj for using Docjure 1.3

    (defproject some.cool/project "1.0.0-SNAPSHOT"
      :description "Spreadsheet magic using Docjure"
      :dependencies [[org.clojure/clojure "1.1.0"]
                     [org.clojure/clojure-contrib "1.1.0"]
                     [dk.ative/docjure "1.3.0"]])


## Installation

You need to install the Leiningen build tool to build the library.
You can get it here: [Leiningen](http://github.com/technomancy/leiningen)

The library uses the Apache POI library which will be downloaded by
the "lein deps" command.

Then build the library:

     lein deps
     lein compile
     lein test


## License

Copyright (c) 2009-2010 Martin Jul, Ative (www.ative.dk)

Docjure is licensed under the MIT License. See the LICENSE file for
the license terms.

Docjure uses the Apache POI library, which is licensed under the
[Apache License v2.0](http://www.apache.org/licenses/LICENSE-2.0).

For more information on Apache POI refer to the
[Apache POI web site](http://poi.apache.org/).




# Docjure

Docjure makes reading and writing Office documents in Clojure easy.

## Usage

### Example: Read a Price List spreadsheet

    (use 'dk.ative.docjure.spreadsheet)       

    ;; Load a spreadsheet and read the first two columns from the 
    ;; price list sheet:
    (->> (load-workbook "spreadsheet.xlsx")
         (select-sheet "Price List")
         (select-columns {:A :name, :B :price}))

    ;=> [{:name "Foo Widget", :price 100}, {:name "Bar Widget", :price 200}]

### Example: Create a spreadsheet 
This example creates a spreadsheet with a single sheet named "Price List".
It has three rows. We apply a style of yellow background colour and bold font
to the top header row, then save the spreadsheet.

    (use 'dk.ative.docjure.spreadsheet)       
    
    ;; Create a spreadsheet and save it
    (let [wb (create-workbook "Price List"
                              [["Name" "Price"]
                               ["Foo Widget" 100]
                               ["Bar Widget" 200]])
          sheet (select-sheet "Price List" wb)
          header-row (first (row-seq sheet))]
      (do
        (set-row-style! header-row (create-cell-style! wb {:background :yellow,
                                                           :font {:bold true}}))
        (save-workbook! "spreadsheet.xlsx" wb)))
    

### Automatically get the Docjure jar from Clojars

The Docjure jar is distributed on [Clojars](http://clojars.org/dk.ative/docjure). 

If you are using the Leiningen build tool just add this line to the
:dependencies list in project.clj to use it:

    [dk.ative/docjure "1.6.0"]	

Remember to issue the 'lein deps' command to download it.

#### Example project.clj for using Docjure 1.6.0

    (defproject some.cool/project "1.0.0-SNAPSHOT"
      :description "Spreadsheet magic using Docjure"
      :dependencies [[org.clojure/clojure "1.4.0"]
                     [dk.ative/docjure "1.6.0"]])


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

Copyright (c) 2009-2014 Martin Jul

Docjure is licensed under the MIT License. See the LICENSE file for
the license terms.

Docjure uses the Apache POI library, which is licensed under the
[Apache License v2.0](http://www.apache.org/licenses/LICENSE-2.0).

For more information on Apache POI refer to the
[Apache POI web site](http://poi.apache.org/).


## Contact information

* [Docjure on GitHub](https://github.com/mjul/docjure)

Martin Jul

* Email: martin@.....com
* Twitter: mjul
* GitHub: [mjul](https://github.com/mjul) 


## Contributors
This library includes great contributions from

* [Carl Baatz](https://github.com/cbaatz) (cbaatz)
* [Michael van Acken](https://github.com/mva) (mva)
* [Ragnar Dahlén](https://github.com/ragnard) (ragnard)
* [Vijay Kiran](https://github.com/vijaykiran) (vijaykiran)
* [Jon Neale](https://github.com/jonneale) (jonneale)
* ["Naipmoro"](https://github.com/naipmoro)(naipmoro)

Thank you very much!

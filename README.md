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


### Example: Handling Error Cells

If the spreadsheet being read contains cells with errors the default
behaviour of the library is to return a keyword representing the
error as the cell value.

For example, given a spreadsheet with errors:

	(use 'dk.ative.docjure.spreadsheet)

	(def sample-cells (->> (load-workbook "spreadsheet.xlsx")
                           (sheet-seq)
                           (mapcat cell-seq)))

    sample-cells

    ;=> (#<XSSFCell 15.0> #<XSSFCell NA()> #<XSSFCell 35.0> #<XSSFCell 13/0> #<XSSFCell 33.0> #<XSSFCell 96.0>)

Reading error cells, or cells that evaluate to an error (e.g. divide by
zero) returns a keyword representing the type of error from
`read-cell`.

	(->> sample-cells
         (map read-cell))

	;=> (15.0 :NA 35.0 :DIV0 33.0 96.0)

How you handle errors will depend on your application. You may want to
replace specific errors with a default value and remove others for
example:

	(->> sample-cells
         (map read-cell)
         (map #(get {:DIV0 0.0} % %))
         (remove keyword?))

	;=> (15.0 35.0 0.0 33.0 96.0)

The following is a list of all possible [error values](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/FormulaError.html#enum_constant_summary):

    #{:VALUE :DIV0 :CIRCULAR_REF :REF :NUM :NULL :FUNCTION_NOT_IMPLEMENTED :NAME :NA}

### Automatically get the Docjure jar from Clojars

The Docjure jar is distributed on [Clojars](http://clojars.org/dk.ative/docjure).

If you are using the Leiningen build tool just add this line to the
:dependencies list in project.clj to use it:

    [dk.ative/docjure "1.9.0"]

Remember to issue the 'lein deps' command to download it.

#### Example project.clj for using Docjure 1.9.0

    (defproject some.cool/project "1.0.0-SNAPSHOT"
      :description "Spreadsheet magic using Docjure"
      :dependencies [[org.clojure/clojure "1.6.0"]
                     [dk.ative/docjure "1.9.0"]])


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

Copyright (c) 2009-2015 Martin Jul

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
* ["Naipmoro"](https://github.com/naipmoro) (naipmoro)
* [Nikolay Durygin](https://github.com/nidu) (nidu)
* [Oliver Holworthy](https://github.com/oholworthy) (oholworthy)
* ["rakhra"](https://github.com/rakhra) (rakhra)
* [Igor Tovstopyat-Nelip](https://github.com/igortn) (igort)
* [Dino Kovač](https://github.com/reisub) (reisub)

Thank you very much!

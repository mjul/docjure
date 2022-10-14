# Docjure

Docjure makes reading and writing Office Excel spreadsheet documents
in Clojure easy.

## Who is this for?
Docjure is aimed at making the basic use case of reading and writing
spreadsheets easy.

If you need advanced charting, pivot tables *etc.*, the easiest way is
to build template spreadsheets with Excel and populate them with
Docjure.

If you want to manipulate advanced features programatically,
you are probably better off using the underlying Apache POI
spreadsheet library directly or looking for another tool.

Docjure has low churn. It is very stable library with a history going
back to 2009 (open-sourced in 2010).


## Usage

### Example: Read a Price List spreadsheet

```clj
(use 'dk.ative.docjure.spreadsheet)

;; Load a spreadsheet and read the first two columns from the
;; price list sheet:
(->> (load-workbook "spreadsheet.xlsx")
     (select-sheet "Price List")
     (select-columns {:A :name, :B :price}))

;=> [{:name "Foo Widget", :price 100}, {:name "Bar Widget", :price 200}]
```

### Example: Read a single cell

If you want to read a single cell value, you can use the `select-cell` function which
takes an Excel-style cell reference (A2) and returns the cell. In order to get the
actual value, use `read-cell`

```clj
(use 'dk.ative.docjure.spreadsheet)
(read-cell
 (->> (load-workbook "spreadsheet.xlsx")
      (select-sheet "Price List")
      (select-cell "A1")))
```

### Example: Load a Workbook from a Resource
This example loads a workbook from a named file. In the case of running
in the application server, the file typically resides in the resources directory,
and it's not on the caller's path. To cover this scenario, we provide
the function 'load-workbook-from-resource' that takes a named resource
as the parameter. After a minor modification, the same example will look like:

```clj
(->> (load-workbook-from-resource "spreadsheet.xlsx")
     (select-sheet "Price List")
     (select-columns {:A :name, :B :price}))
```

### Example: Load a Workbook from a Stream
The function 'load-workbook' is a multimethod, and the first example takes
a file name as a parameter. The overloaded version of 'load-workbook'
takes an InputStream. This may be useful when uploading a workbook to the server
over HTTP connection as multipart form data. In this case, the web framework
passes a byte buffer, and the example should be modified as (note that you have
to use 'with-open' to ensure that the stream will be closed):

```clj

(with-open [stream (clojure.java.io/input-stream bytes)]
  (->> (load-workbook stream)
       (select-sheet "Price List")
       (select-columns {:A :name, :B :price})))
```

### Example: Create a spreadsheet
This example creates a spreadsheet with a single sheet named "Price List".
It has three rows. We apply a style of yellow background colour and bold font
to the top header row, then save the spreadsheet.

```clj
(use 'dk.ative.docjure.spreadsheet)

;; Create a spreadsheet and save it
(let [wb (create-workbook "Price List"
                          [["Name" "Price"]
                           ["Foo Widget" 100]
                           ["Bar Widget" 200]])
      sheet (select-sheet "Price List" wb)
      header-row (first (row-seq sheet))]
  (set-row-style! header-row (create-cell-style! wb {:background :yellow,
                                                     :font {:bold true}}))
  (save-workbook! "spreadsheet.xlsx" wb))
```

### Example: Create a workbook with multiple sheets
This example creates a spreadsheet with multiple sheets. Simply add more
sheet-name and data pairs. To create a sheet with no data, pass `nil` as
the data argument.

```clj
(use 'dk.ative.docjure.spreadsheet)

;; Create a spreadsheet and save it
(let [wb (create-workbook "Squares"
                          [["N" "N^2"]
                           [1 1]
                           [2 4]
                           [3 9]]
                          "Cubes"
                          [["N" "N^3"]
                           [1 1]
                           [2 8]
                           [3 27]])]
   (save-workbook! "exponents.xlsx" wb))
```

### Example: Use Excel Formulas in Clojure

Docjure allows you not only to evaluate a formula cell in a speadsheet, it also
provides a way of exposing a formula in a cell as a Clojure function using the
`cell-fn` function.

    (use 'dk.ative.docjure.spreadsheet)
    ;; Load a speadsheet and take the first sheet, construct a function from cell A2, taking
    ;; A1 as input.
    (def formula-from-a2 (cell-fn "A2"
                                      (first (sheet-seq (load-workbook "spreadsheet.xlsx")))
                                      "A1"))

    ;; Returns value of cell A2, as if value in cell A1 were 1.0
    (formula-from-a2 1.0)

### Example: Handling Error Cells

If the spreadsheet being read contains cells with errors the default
behaviour of the library is to return a keyword representing the
error as the cell value.

For example, given a spreadsheet with errors:

```clj
(use 'dk.ative.docjure.spreadsheet)

(def sample-cells (->> (load-workbook "spreadsheet.xlsx")
                       (sheet-seq)
                       (mapcat cell-seq)))

sample-cells

;=> (#<XSSFCell 15.0> #<XSSFCell NA()> #<XSSFCell 35.0> #<XSSFCell 13/0> #<XSSFCell 33.0> #<XSSFCell 96.0>)
```

Reading error cells, or cells that evaluate to an error (e.g. divide by
zero) returns a keyword representing the type of error from
`read-cell`.

```clj
(->> sample-cells
     (map read-cell))

;=> (15.0 :NA 35.0 :DIV0 33.0 96.0)
```

How you handle errors will depend on your application. You may want to
replace specific errors with a default value and remove others for
example:

```clj
(->> sample-cells
     (map read-cell)
     (replace {:DIV0 0.0})
     (remove keyword?))

;=> (15.0 35.0 0.0 33.0 96.0)
```

The following is a list of all possible [error values](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/FormulaError.html#enum_constant_summary):

```clj
#{:VALUE :DIV0 :CIRCULAR_REF :REF :NUM :NULL :FUNCTION_NOT_IMPLEMENTED :NAME :NA}
```

### Example: Iterating over spreadsheet data

#### A note on sparse data

It's worth understanding a bit about the underlying structure of a spreadsheet before you
start iterating over the contents.

Spreadsheets are designed to be sparse - not all rows in the spreadsheet must physically exist,
and not all cells in a row must physically exist.  This is how you can create data at ZZ:65535 without
using huge amounts of storage.

Thus each cell can be in 3 states - with data, blank, or nonexistent (null).  There's a special type [CellType.BLANK](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/CellType.html#BLANK) for blank cells, but missing cells are just returned as nil.

Similarly rows can exist with cells, or exist but be empty, or they can not exist at all.

Prior to Docjure 1.11 the iteration functions wrapped the underlying Apache POI iterators, which skipped over missing data - this could cause surprising behaviour, especially when there were missing cells inside tabular data.

Since Docjure 1.11 iteration now returns `nil` values for missing rows and cells - this is a *breaking change* - any code that calls `row-seq` or `cell-seq` now needs to deal with possible nil values.

#### Iterating over rows

You can iterate over all the rows in a worksheet with `row-seq`:

```clj
(->> (load-workbook "test.xls")
     (select-sheet "first")
     row-seq)
```

This will return a sequence of `org.apache.poi.usermodel.Row` objects, or `nil` for any missing rows.  You can use `(remove nil? (row-seq ...) )` if you are happy to ignore missing rows, but then be aware the nth result in the sequence might not match the nth row in the spreadsheet.


#### Iterating over cells

You can iterate over all the cells in a row with `cell-seq` - this returns a sequence of `org.apache.poi.usermodel.Cell` objects, or `nil` for missing cells.  Note that `(read-cell nil)` returns `nil` so it's safe to apply `read-cell` to the results of `cell-seq`

```clj
(->> (load-workbook "test.xls")
     (select-sheet "first")
     row-seq
     (remove nil?)
     (map cell-seq)
     (map #(map read-cell %)))
```

For example, if you run the above snippet on a sparse spreadsheet like:

| First Name | Middle Name | Last Name |
| --- | --- | --- |
| Edger | Allen | Poe |
| `(missing row)` |
| John | `(missing)` | Smith |

Then it will return:

```clj
(("First Name" "Middle Name" "Last Name")
 ("Edger" "Allen" "Poe")
 ("John" nil "Smith"))
```

### Formatting: Adjust Column Width to Contents

You can adjust the column widths to they fit the contents.
This makes generated workbooks look nicer.

For example, to auto-size all the columns in all the sheets
in a workbook, use this:

```clj
;; wb is a workbook
(dorun (for [sheet (sheet-seq wb)]
             (auto-size-all-columns! sheet)))
```

To apply auto-width to individual columns in a sheet, use the
`auto-size-column!` function.



### Automatically get the Docjure jar from Clojars

The Docjure jar is distributed on
[Clojars](http://clojars.org/dk.ative/docjure). Here you can find both
release builds and snapshot builds of pre-release versions.

If you are using the Leiningen build tool just add this line to the
:dependencies list in project.clj to use it:

```clj
[dk.ative/docjure "1.14.0"]
```

Remember to issue the 'lein deps' command to download it.



#### Example project.clj for using Docjure 1.17.0

```clj
(defproject some.cool/project "1.0.0-SNAPSHOT"
      :description "Spreadsheet magic using Docjure"
      :dependencies [[org.clojure/clojure "1.10.0"]
                     [dk.ative/docjure "1.17.0"]])
```

## Installation for Contributors
You need to install the Leiningen build tool to build the library.
You can get it here: [Leiningen](http://github.com/technomancy/leiningen)

The library uses the Apache POI library which will be downloaded by
the "lein deps" command.

Then build the library:

     lein deps
     lein compile
     lein test

To run the tests on all supported Clojure versions use:

    lein all test

To check for security issues use:

    lein nvd check

To check for new versions of dependencies:

    lein ancient

### Releasing to Clojars
When releasing a version to Clojars you must provide your user-name. The password is a deployment token, not your normal password. You can generate this by logging into Clojars. These tokens have an expiration date so if it does not work, log in a check if you need a new token.


```
    lein deploy clojars
```

You also need a GPG key to sign the releases.

These also expire, and must be periodically renewed. You can check your keys and their status like this:


```
    gpg --list-keys
```

See also `lein help deploy` and `lein help gpg`.

Remember to tag releases in git. You can list the tags with `git tag -n` and tag with `git tag -a TAGNAME -m 'Tag comment'`.


## Build Status
[![Build Status](https://travis-ci.org/mjul/docjure.svg?branch=master)](https://travis-ci.org/mjul/docjure)

## License

Copyright (c) 2009-2022 Martin Jul

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
* [Igor Tovstopyat-Nelip](https://github.com/igortn) (igortn)
* [Dino Kovač](https://github.com/reisub) (reisub)
* [Lars Trieloff](https://github.com/trieloff) (trieloff)
* [Jens Bendisposto](https://github.com/bendisposto) (bendisposto)
* [Stuart Hinson](https://github.com/stuarth) (stuarth)
* [Dan Petranek](https://github.com/dpetranek) (dpetranek)
* [Aleksander Madland Stapnes](https://github.com/madstap) (madstap)
* [Korny Sietsma](https://github.com/kornysietsma) (kornysietsma)
* [Antti Virtanen](https://github.com/lokori) (lokori)
* [alephyud](https://github.com/alephyud) (alephyud)
* [Markku Rontu](https://github.com/Macroz) (macroz)
* [Harold](https://github.com/harold) (harold)
* [Alex Scott](https://github.com/axrs) (axrs)
* [Maurício Szabo](https://github.com/mauricioszabo) (mauricioszabo)
* [Mic Sokoli](https://github.com/MicSokoli) (MicSokoli)
* [Ekaitz Zárraga](https://github.com/ekaitz-zarraga) (ekaitz-zarraga)
* [Ross Gibb](https://github.com/rossgibb) (rossgibb)
* [Manuel Herzog](https://github.com/manuelherzog) (manuelherzog)
* [Francisco Vides Fernández](https://github.com/fvides) (fvides)
* [WonJun Lee](https://github.com/Lee-WonJun) (Lee-WonJun)
* [Jari Hanhela](https://github.com/Jarzka) (Jarzka)
* [Maxim Penzin](https://github.com/maxp) (maxp)
* [Christopher Miles](https://github.com/cmiles74) (cmiles74)
* [Danny Freeman](https://github.com/dannyfreeman) (dannyfreeman)
* [Orestis Markou](https://github.com/orestis) (orestis)

Thank you very much!

## Honorary Mention

A special thank you also goes out to people that did not contribute code
but shared their ideas, reported security issues or bugs and otherwise
inspired the continuing work on the project.


* [vemv](https://github.com/vemv) (vemv)

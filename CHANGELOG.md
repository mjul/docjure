# CHANGELOG for Docjure

## Version 1.16.0
* Upgraded to use Apache POI 4.1.1 (fixes for CVE-2019-12415)

## Version 1.15.0
* Upgraded to use Clojure 1.10.1 (from 1.10.0)

## Version 1.14.0
* Added a font cache to create-date-format to prevent overflowing the number
  of styles in the document by reusing the styles.
* Fixed some documentation errors
* Dropped support for Java SE 6 and 7
* (Java SE 8 is still supported)
* Added support for Java SE 10, 11, 12 and 13
* Added `escape-cell` utility function to escape text clashing with Excel's `_x..._` Unicode notation.
* Upgraded to use Clojure 1.10 (test suite runs on Clojure 1.5 through 1.10)
* Upgraded to use Apache POI 4.1.0

## Version 1.13.0
* Dropped support for Clojure 1.3 and 1.4.
* Upgraded to use Clojure 1.9 (test suite runs on Clojure 1.5 through 1.9)

## Version 1.12.0
* Improved documentation.
* Added more cell style formatting options.
* Upgraded to Apache POI v3.17.

## Version 1.11.0
* Upgraded to Apache POI v3.14.
* Added support for sparse data in seq functions. Previously the reader
would skip blank row, now they will be return as `nil`. We consider
this a bug-fix, so there is no major version update for this. However,
note that this is potentially a *breaking change* if you use `row-seq` or `cell-seq` or
related functions, and you relay on blank missing rows/cells in your
spreadsheets to be ignored.
* Added support for multi-sheet workbooks.
* Added support for fomulae.

## Version 1.10.0
* Upgraded to Apache POI v3.13.
* Add `select-cell` function to easily read a single cell value.
* Upgraded to use Clojure 1.8 (test suite runs on Clojure 1.3 through 1.8)

## Version 1.9.0

* `read-cell` now works on error cells and non-numeric formula cells without throwing an exception. (All cell types now handled safely).

Error cells return keyword of the error type:

```
:VALUE :DIV0 :CIRCULAR_REF :REF :NUM :NULL :FUNCTION_NOT_IMPLEMENTED :NAME :NA
```

* Added functions to load workbooks from streams and resources in
  addition to files: `load-workbook-from-stream`,
  `load-workbook-from-resource` and `load-workbook-from-file`.
* Make `load-workbook` a multi-method accepting a string (filename) or a
  stream.

## Version 1.8.0
* Upgraded to use Clojure 1.6 as default Clojure version.
* Upgraded to Apache POI v3.11.
* Now handles both 1900 and 1904-based dates for the Mac and Windows
 versions of Excel [More info](http://support.microsoft.com/kb/180162).

## Version 1.7.0
* `select-sheet` can now select sheets by regex and predicate functions in addition to exact sheet name (contributed by jonneale).
* upgraded Clojure version to 1.5.1
* added font and cell styling options (contributed by naipmoro):
* added option work on the legacy Excel 'XLS' file format (`create-xls-workbook`)
* added font styling options to `create-font!`:

    :name (font family - string)
    :size (font size - integer)
    :color (font color - keyword)
    :bold (true | false)
    :italic (true | false)
    :underline (true | false)

* added styling options to `create-cell-style!`:

    :background (background colour - keyword)
    :font (font | fontmap of font options)
    :halign (:left | :right | :center)
    :valign (:top | :bottom | :center)
    :wrap (true | false - controls text wrapping)
    :border-left (:thin | :medium | :thick)
    :border-right (:thin | :medium | :thick)
    :border-top (:thin | :medium | :thick)
    :border-bottom (:thin | :medium | :thick)



## Version 1.6.0
* Upgraded to Clojure 1.4 as the default Clojure version.
* Upgraded to Apache POI 3.9
* Support for Clojure 1.3, 1.4 and 1.5 (RC 1) via lein profiles (contributed by ragnard)
* Support for Travis-CI (contributed by ragnard)
* Use type hints to call correct overload for setting nil date (contributed by mva)

## Version 1.5.0
* Introduces remove-row! and remove-all-rows!.
* Adds row-vec function to create row data for adding to sheet from a struct to ease writing select, transform, write-back tasks.
* Formulas are now evaluated when they are read from the sheet and the resulting value is returned.
* Added named ranges functions add-name! and select-name (contributed by cbaatz).
* Added row style functions set-row-style! and get-row-styles for styling rows (contributed by cbaatz).

## Version 1.4
* Introduces cell styling (font control, background colour).
* A more flexible cell-seq (supports sheet, row or collections of these).

## Version 1.3
* Updated semantics for reading blank cells: now they are read as nil (formerly read as empty strings).

## Version 1.2

First public release.

## Earlier versions

Earlier versions used internally for projects in Ative in 2009 and 2010.

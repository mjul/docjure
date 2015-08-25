# CHANGELOG for Docjure

## Version 1.10.0
* Upgraded to Apache POI v3.12.

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
* select-sheet can now select sheets by regex and predicate functions in addition to exact sheet name (contributed by jonneale).
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

# CHANGELOG for Docjure

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




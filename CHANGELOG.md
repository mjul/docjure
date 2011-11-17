# CHANGELOG for Docjure

## Version 1.5.0-SNAPSHOT
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

Earlier versions used internally for projects in Ative in 2009 and
2010.




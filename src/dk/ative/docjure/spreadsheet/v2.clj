(ns dk.ative.docjure.spreadsheet.v2
  (:require [dk.ative.docjure.spreadsheet :as spreadsheet])
  (:import (org.apache.poi.ss.usermodel Workbook Sheet Cell Row)))

(defprotocol Context
  (select-sheet [context predicate])
  (select-cell [context ref])
  (read-cell [context])

  (sheet-seq [context])
  (row-seq [context])
  (cell-seq [context])

  (add-style! [context style])

  (set-cell-style! [context])
  (set-row-styles! [context]))

(defrecord POIContext [^Workbook workbook ^Sheet sheet ^Row row ^Cell cell styles options]
  Context
  (select-sheet [context predicate]
    (assert (instance? Workbook workbook) "We need a workbook to be able to select a sheet")
    (assoc context :sheet (spreadsheet/select-sheet predicate workbook)))
  (select-cell [context ref]
    (assert (instance? Sheet sheet) "We require a sheet to be able to select a cell, please use select-sheet")
    (assoc context :cell (spreadsheet/select-cell ref sheet)))
  (read-cell [_]
    (assert (instance? Cell cell) "Please select a cell using select-cell")
    (spreadsheet/read-cell cell options))

  (sheet-seq [context]
    (assert (instance? Workbook workbook) "We need a workbook to be able to select a sheet")
    (map #(assoc context :sheet %) (spreadsheet/sheet-seq workbook)))
  (row-seq [context]
    (assert (instance? Sheet sheet) "")
    (map #(assoc context :row %) (spreadsheet/row-seq sheet)))
  (cell-seq [context]
    (assert (or (instance? Sheet sheet)
                (instance? Row row)) "")
    (map #(assoc context :cell %) (spreadsheet/cell-seq (or row sheet))))

  (add-style! [context style]
    (assert (instance? Workbook workbook) "We need a workbook to be able to select a sheet")
    (update context :styles #(conj (or % []) (spreadsheet/create-cell-style! workbook style))))

  (set-cell-style! [context]
    (assert (instance? Cell cell) "Please select a cell using select-cell")
    (spreadsheet/set-cell-style! cell (first styles)))
  (set-row-styles! [context]
    (assert (instance? Row row) "")
    (spreadsheet/set-row-styles! row styles)))

(defn load-workbook [input & [options]]
  (map->POIContext {:workbook (spreadsheet/load-workbook input)
                    :options (or options {})}))

(defn create-workbook [sheet-name data & [options]]
  (map->POIContext {:workbook (spreadsheet/create-workbook sheet-name data)
                    :options (or options {})}))

(defmacro with-styles [context styles & body]
  `(let [~context (reduce add-style! ~context ~styles)]
     ~@body))

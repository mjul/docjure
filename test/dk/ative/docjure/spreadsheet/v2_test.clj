(ns dk.ative.docjure.spreadsheet.v2-test
  (:require [clojure.test :refer :all]
            [dk.ative.docjure.spreadsheet.v2 :refer :all])
  (:import (org.apache.poi.ss.usermodel IndexedColors)))

(def config {:simple "test/dk/ative/docjure/testdata/simple.xlsx"
             :missing-workbook "test/dk/ative/docjure/testdata/missing-workbook.xlsx"})

(deftest load-and-read-simple
  (let [workbook (load-workbook (:simple config))
        sheet (first (sheet-seq workbook))]
    (is (= 1.0 (read-cell (select-cell sheet "A2"))))))

(deftest missing-workbooks-causes-explosions
  (let [workbook (load-workbook (:missing-workbook config))
        sheet (first (sheet-seq workbook))]
    (is (thrown? java.lang.RuntimeException
                 (read-cell (select-cell sheet "A1"))))))

(deftest ignore-missing-workbooks-uses-cached-value
  (let [workbook (load-workbook (:missing-workbook config)
                                {:evaluator {:ignore-missing-workbooks? true}})
        sheet (first (sheet-seq workbook))]
    (is (= 6.0 (read-cell (select-cell sheet "A1"))))))

(deftest formatting-cells
  (let [workbook (load-workbook (:simple config))
        sheet (first (sheet-seq workbook))]
    (with-styles sheet [{:background :yellow}]
      (let [cell (select-cell sheet "A1")]
        (set-cell-style! cell)
        (is (= (.getIndex IndexedColors/YELLOW) (.. (:cell cell) getCellStyle getFillForegroundColor)))))))

(deftest formatting-rows
  (let [workbook (create-workbook "Dummy" [["foo" "bar"] ["data b" "data b"]])
        [header-row data-row] (row-seq (select-sheet workbook "Dummy"))
        [a1 b1] (cell-seq header-row)
        [a2 b2] (cell-seq data-row)]
    (with-styles header-row [{:background :yellow} {:background :red}]
      (set-row-styles! header-row))
    (is (= (.getIndex IndexedColors/YELLOW) (.. (:cell a1) getCellStyle getFillForegroundColor)))
    (is (= (.getIndex IndexedColors/RED) (.. (:cell b1) getCellStyle getFillForegroundColor)))
    (is (not= (.getIndex IndexedColors/YELLOW) (.. (:cell a2) getCellStyle getFillForegroundColor)))
    (is (not= (.getIndex IndexedColors/RED) (.. (:cell b2) getCellStyle getFillForegroundColor)))))

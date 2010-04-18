(ns docjure.spreadsheet-test
  (:use [docjure.spreadsheet] :reload-all)
  (:use [clojure.test])
  (:import (org.apache.poi.ss.usermodel Workbook Sheet Cell Row)
	   (org.apache.poi.xssf.usermodel XSSFWorkbook)))

(def config {:datatypes-file "test/docjure/testdata/datatypes.xlsx"})
(def datatypes-map {:A :text, :B :integer, :C :decimal, :D :date, :E :time, :F :date-time, :G :percentage, :H :fraction, :I :scientific})

(defn- datatypes-rows [file]
  (->> (load-workbook file) 
       sheet-seq
       first
       (select-columns datatypes-map)))

(defn- datatypes-data [file column]
  (->> (datatypes-rows file)
       rest
       (map column)
       (remove nil?)))

(deftest add-sheet!-test
  (let [workbook (XSSFWorkbook.)
	sheet-name "Sheet 1"
	actual   (add-sheet! workbook sheet-name)]
    (testing "Sheet creation"
      (is (= 1 (.getNumberOfSheets workbook)) "Expected sheet to be added.")
      (is (= sheet-name (.. workbook (getSheetAt 0) (getSheetName))) "Expected sheet to have correct name."))))

(deftest create-workbook-test
  (let [sheet-name "Sheet 1" 
	sheet-data [["A1" "B1" "C1"]
		    ["A2" "B2" "C2"]]
	workbook (create-workbook sheet-name sheet-data)]
    (testing "Sheet creation"
      (is (= 1 (.getNumberOfSheets workbook)) "Expected sheet to be added.")
      (is (= sheet-name (.. workbook (getSheetAt 0) (getSheetName))) "Expected sheet to have correct name."))
    (testing "Sheet data"
      (let [sheet (.getSheetAt workbook 0)
	    rows  (vec (iterator-seq (.iterator sheet)))]
	(is (= (count sheet-data) (.getPhysicalNumberOfRows sheet)) "Expected correct number of rows.")
	(is (= (count (first sheet-data)) (.getLastCellNum (first rows))) "Expected correct number of columns.")
	(are [actual-cell expected-value] (= expected-value (.getStringCellValue actual-cell))
	     (.getCell (first rows) 0) (ffirst sheet-data)
	     (.getCell (first rows) 1) (second (first sheet-data))
	     (.getCell (second rows) 0) (first (second sheet-data))
	     (.getCell (second rows) 1) (second (second sheet-data)))))))

(deftest sheet-seq-test
  (let [sheet-name "Sheet 1" 
	sheet-data [["foo" "bar"]]
	workbook (create-workbook sheet-name sheet-data)
	actual   (sheet-seq workbook)]
    (is (not (nil? actual)))
    (is (= 1 (count actual)))
    (is (= sheet-name (.getSheetName (first actual))))))


(deftest select-columns-test
  (let [data     [["Name" "Quantity" "Price"] 
		  ["foo" 1 42] 
		  ["bar" 2 108]]
	workbook (create-workbook "Sheet 1" data)
	sheet    (first (sheet-seq workbook))]
    (testing "Find existing columns should create map."
      (let [rows (select-columns {:A :name, :B :quantity} sheet)]
	(is (= (count data) (count rows)))
	(is (every? #(= 2 (count (keys %))) rows))
	(is (every? #(and (contains? % :name)
			  (contains? % :quantity)) rows))
	(are [actual expected] (= actual (zipmap [:name :quantity] expected))
	     (first rows) (data 0)
	     (second rows) (data 1)
	     (nth rows 2) (data 2))))
    (testing "Find non-existing columns should map to empty maps."
      (let [rows (select-columns {:X :colX, :Y :colY} sheet)]
	(is (= (count data) (count rows)))
	(is (every? empty? rows))))))


(deftest load-workbook-integration-test
  (let [file (config :datatypes-file)
	loaded (load-workbook file)]
    (is (not (nil? loaded))
    (is (isa? (class loaded) Workbook)))))

 
(deftest select-columns-integration-test
  (testing "Reading text fields."
    (let [file (config :datatypes-file)]
      (is (every? string? (datatypes-data file :text)))
;      (is (every? integer? (datatypes-data file :integer)))
;      (is (every? decimal? (datatypes-data file :decimal)))
      )))


(ns dk.ative.docjure.spreadsheet-test
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:use [clojure.test])
  (:import (org.apache.poi.ss.usermodel Workbook Sheet Cell Row)
	   (org.apache.poi.xssf.usermodel XSSFWorkbook)
	   (java.util Date)))

(def config {:datatypes-file "test/dk/ative/docjure/testdata/datatypes.xlsx"})
(def datatypes-map {:A :text, :B :integer, :C :decimal, :D :date, :E :time, :F :date-time, :G :percentage, :H :fraction, :I :scientific})

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
	(is (= 0 (.getRowNum (first rows))) "Expected correct row number.")
	(is (= (count (first sheet-data)) (.getLastCellNum (first rows))) "Expected correct number of columns.")
	(are [actual-cell expected-value] (= expected-value (.getStringCellValue actual-cell))
	     (.getCell (first rows) 0) (ffirst sheet-data)
	     (.getCell (first rows) 1) (second (first sheet-data))
	     (.getCell (second rows) 0) (first (second sheet-data))
	     (.getCell (second rows) 1) (second (second sheet-data)))))))


(deftest sheet-seq-test
  (let [sheet-name "Sheet 1" 
	sheet-data [["foo" "bar"]]]
    (testing "Empty workbook"
      (let [workbook (XSSFWorkbook.)
	    actual (sheet-seq workbook)]
	(is (not (nil? actual)))
	(is (empty? actual))))
    (testing "Single sheet."
      (let [workbook (create-workbook sheet-name sheet-data)
	    actual   (sheet-seq workbook)]
	(is (= 1 (count actual)))
	(is (= sheet-name (.getSheetName (first actual))))))
    (testing "Multiple sheets."
      (let [workbook (create-workbook sheet-name sheet-data)
	    sheet2 (.createSheet workbook "Sheet 2")
	    sheet3 (.createSheet workbook "Sheet 3")
	    actual (sheet-seq workbook)]
	(is (= 3 (count actual)))
	(is (= [sheet2 sheet3] (rest actual)))))))


(deftest sheet-name-test
  (let [name       "Sheet 1" 
	data       [["foo" "bar"]]
	workbook   (create-workbook name data)
	sheet      (first (sheet-seq workbook))]
    (is (= name (sheet-name sheet)) "Expected correct sheet name.")))

(deftest select-sheet-test
  (let [name       "Sheet 1" 
	data       [["foo" "bar"]]
	workbook   (create-workbook name data)
	first-sheet (first (sheet-seq workbook))]
    (is (= first-sheet (select-sheet name workbook)) "Expected to find the sheet.")
    (is (nil? (select-sheet "unknown name" workbook)) "Expected to get nil for no match.")))

(deftest select-columns-test
  (let [data     [["Name" "Quantity" "Price" "On Sale"] 
		  ["foo" 1 42 true] 
		  ["bar" 2 108 false]]
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
	(is (every? empty? rows))))
    (testing "Should support many datatypes."
      (let [rows (select-columns {:A :string, :B :number, :D :boolean} sheet)
	    data-rows (rest rows)]
	(are [actual expected] (= actual (let [[a b c d] expected] 
					   {:string a, :number b, :boolean d}))
	     (first data-rows) (data 1)
	     (second data-rows) (data 2))))))


(deftest load-workbook-integration-test
  (let [file (config :datatypes-file)
	loaded (load-workbook file)]
    (is (not (nil? loaded))
    (is (isa? (class loaded) Workbook)))))


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

(defn- date? [date] 
  (isa? (class date) Date))

(deftest select-columns-integration-test
  (testing "Reading text fields."
    (let [file (config :datatypes-file)]
      (is (every? string? (datatypes-data file :text)))
      (is (every? number? (datatypes-data file :integer)))
      (is (every? number? (datatypes-data file :decimal)))
      (is (every? date? (datatypes-data file :date)))
      (is (every? date? (datatypes-data file :time)))
      (is (every? date? (datatypes-data file :date-time)))
      (is (every? number? (datatypes-data file :percentage)))
      (is (every? number? (datatypes-data file :fraction)))
      (is (every? number? (datatypes-data file :scientific))))))



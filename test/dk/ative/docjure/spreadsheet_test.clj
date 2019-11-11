(ns dk.ative.docjure.spreadsheet-test
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:use [clojure.test])
  (:require [clojure.java.io :as io])
  (:import (org.apache.poi.ss.usermodel Workbook Row
                                        Row$MissingCellPolicy
                                        CellStyle IndexedColors Font
                                        CellValue HorizontalAlignment
                                        VerticalAlignment BorderStyle
                                        FillPatternType)
           (org.apache.poi.xssf.usermodel XSSFWorkbook XSSFFont)
           (java.util Date)
           (java.io FileInputStream)))

(def config {:datatypes-file  "test/dk/ative/docjure/testdata/datatypes.xlsx"
             :datatypes-test-resource "dk/ative/docjure/testdata/datatypes.xlsx"
             :formulae-file   "test/dk/ative/docjure/testdata/formulae.xlsx"
             :blanks-file "test/dk/ative/docjure/testdata/blanks.xlsx"
             :1900-based-file "test/dk/ative/docjure/testdata/1900-based-dates.xlsx"
             :1904-based-file "test/dk/ative/docjure/testdata/1904-based-dates.xlsx"
             :simple "test/dk/ative/docjure/testdata/simple.xlsx"
             :save-workbook-location "test/dk/ative/docjure/testdata/saved.xlsx"})

(def datatypes-map {:A :text, :B :integer, :C :decimal, :D :date, :E :time, :F :date-time, :G :percentage, :H :fraction, :I :scientific, :J :date-formulae})
(def formulae-map {:A :formula, :B :expected})

(deftest add-sheet!-test
  (let [workbook (XSSFWorkbook.)
        sheet-name "Sheet 1"
        actual   (add-sheet! workbook sheet-name)]
    (testing "Sheet creation"
      (is (= 1 (.getNumberOfSheets workbook)) "Expected sheet to be added.")
      (is (= sheet-name (.. workbook (getSheetAt 0) (getSheetName))) "Expected sheet to have correct name."))
    (testing "Should fail on non-Workbook"
      (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (add-sheet! "not-a-workbook" "sheet-name"))))))

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

(deftest create-workbook-with-multiple-sheets-test
  (let [sheet-1-name "Sheet 1"
        sheet-1-data [["A1" "B1" "C1"]
                      ["A2" "B2" "C2"]]
        sheet-2-name "Sheet 2"
        sheet-2-data [["A1" "B1" "C1"]
                      ["A2" "B2" "C2"]]
        workbook (create-workbook sheet-1-name sheet-1-data sheet-2-name sheet-2-data)]
    (testing "Multiple sheet creation"
      (is (= 2 (.getNumberOfSheets workbook)) "Expected 2 sheets to be added.")
      (is (= sheet-1-name (.. workbook (getSheetAt 0) (getSheetName)))
          "Expected sheet 1 to have correct name.")
      (is (= sheet-2-name (.. workbook (getSheetAt 1) (getSheetName)))
          "Expected sheet 2 to have correct name."))
    (testing "Sheet data"
      (let [sheet-1 (.getSheetAt workbook 0)
            sheet-1-rows (vec (iterator-seq (.iterator sheet-1)))
            sheet-2 (.getSheetAt workbook 1)
            sheet-2-rows (vec (iterator-seq (.iterator sheet-2)))]
        (is (= (count sheet-1-data) (.getPhysicalNumberOfRows sheet-1))
            "Expected correct number of rows for sheet 1.")
        (is (= (count sheet-2-data) (.getPhysicalNumberOfRows sheet-2))
            "Expected correct number of rows for sheet 2.")
        (is (= 0 (.getRowNum (first sheet-1-rows))) "Expected correct row number for sheet 1.")
        (is (= 0 (.getRowNum (first sheet-2-rows))) "Expected correct row number for sheet 2.")
        (is (= (count (first sheet-1-data)) (.getLastCellNum (first sheet-1-rows)))
            "Expected correct number of columns for sheet 1.")
        (is (= (count (first sheet-2-data)) (.getLastCellNum (first sheet-2-rows)))
            "Expected correct number of columns for sheet 2.")
        (are [actual-cell expected-value] (= expected-value (.getStringCellValue actual-cell))
              (.getCell (first sheet-1-rows) 0) (ffirst sheet-1-data)
              (.getCell (first sheet-1-rows) 1) (second (first sheet-1-data))
              (.getCell (second sheet-1-rows) 0) (first (second sheet-1-data))
              (.getCell (second sheet-1-rows) 1) (second (second sheet-1-data))

              (.getCell (first sheet-2-rows) 0) (ffirst sheet-2-data)
              (.getCell (first sheet-2-rows) 1) (second (first sheet-2-data))
              (.getCell (second sheet-2-rows) 0) (first (second sheet-2-data))
              (.getCell (second sheet-2-rows) 1) (second (second sheet-2-data)))))))

(deftest create-sparse-workbook-test
  (let [sheet-name "Sheet 1"
        sheet-data [nil
                    ["A2" nil "C2" "D2"]
                    nil
                    nil
                    [nil "B5" nil "D5" nil]
                    nil]
        workbook (create-sparse-workbook sheet-name sheet-data)]
    (testing "Sparse Sheet creation"
      (is (= 1 (.getNumberOfSheets workbook)) "Expected sheet to be added.")
      (is (= sheet-name (.. workbook (getSheetAt 0) (getSheetName))) "Expected sheet to have correct name."))
    (testing "Sparse Sheet data"
      (let [sheet (.getSheetAt workbook 0)
            rows  (map #(.getRow sheet %)
                       (range 0 (inc (.getLastRowNum sheet))))
            cellvalue (fn [r c] (as-> sheet x
                                    (.getRow x r)
                                    (.getCell x c)
                                    (when x (.getStringCellValue x))))]
        (is (= (dec (count sheet-data)) (inc (.getLastRowNum sheet))) "Expected correct number of rows")
        (is (= (count (filter identity sheet-data)) (.getPhysicalNumberOfRows sheet)) "Expected correct number of physical rows.")
        (is (= [true false true true false] (map nil? rows)) "Expected not all rows created")
        (is (= 1 (.getRowNum (second rows))) "Expected correct row number.")

        (is (= (count (second sheet-data)) (.getLastCellNum (second rows))) "Expected correct number of columns.")
        (is (= 4 (.getLastCellNum (nth rows 4))) "Expected correct number of columns.")
        (is (= "A2" (cellvalue 1 0)))
        (is (= nil (cellvalue 1 1)))
        (is (= "C2" (cellvalue 1 2)))
        (is (= "D2" (cellvalue 1 3)))
        (is (= nil (cellvalue 4 0)))
        (is (= "B5" (cellvalue 4 1)))
        (is (= "D5" (cellvalue 4 3)))))))
        

(deftest row-vec-test
  (testing "Should transform row struct to row vector."
    (is (= ["foo" "bar"] (row-vec [:foo :bar] {:foo "foo", :bar "bar"}))
     "Should map all columns.")
    (is (= ["bar" "foo"] (row-vec [:bar :foo] {:foo "foo", :bar "bar"}))
     "Should respect column order.")
    (is (= [nil nil] (row-vec [:foo :bar] {})) "Should generate all columns.")
    (is (= [] (row-vec [] {:foo "foo", :bar "bar"})) "Should accept empty column-order.")))

(deftest add-row!-test
  (testing "Should fail on invalid parameter types."
    (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (add-row! "not-a-sheet" [1 2 3])))))

(deftest add-rows!-test
  (testing "Should fail on invalid parameter types."
    (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (add-rows! "not-a-sheet" [[1 2 3]])))))

(deftest remove-row!-test
  (let [sheet-name "Sheet 1"
        sheet-data [["A1" "B1" "C1"]
                    ["A2" "B2" "C2"]]
        workbook (create-workbook sheet-name sheet-data)
        sheet (select-sheet sheet-name workbook)
        first-row (first (row-seq sheet))]
    (testing "Should fail on invalid parameter types."
      (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (remove-row! "not-a-sheet" (first (row-seq sheet)))))
      (is (thrown-with-msg? IllegalArgumentException #"row.*" (remove-row! sheet "not-a-row"))))
    (testing "Should remove row."
      (do
       (is (= sheet (remove-row! sheet first-row)))
       (is (= 1 (.getPhysicalNumberOfRows sheet)))
       (is (= [{:A "A2", :B "B2", :C "C2"}] (select-columns {:A :A, :B :B :C :C} sheet)))))))

(deftest remove-all-row!-test
  (let [sheet-name "Sheet 1"
        sheet-data [["A1" "B1" "C1"]
                    ["A2" "B2" "C2"]]
        workbook (create-workbook sheet-name sheet-data)
        sheet (first (sheet-seq workbook))]
    (testing "Should remove all rows."
      (do
       (is (= sheet (remove-all-rows! sheet)))
       (is (= 0 (.getPhysicalNumberOfRows sheet)))))
    (testing "Should fail on invalid parameter types."
      (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (remove-all-rows! "not-a-sheet"))))))

(defn date [year month day]
  (Date. (- year 1900) (dec month) day))

(defn july [day]
  (Date. 2010 7 day))

(deftest read-cell-value-test
    (let [date (july 1)
          workbook (create-workbook "Just a date" [[date]])
          sheet (.getSheetAt workbook 0)
          rows  (vec (iterator-seq (.iterator sheet)))
          data-row (vec (iterator-seq (.cellIterator (first rows))))
          date-cell (first data-row)]
      (testing "Should read all cell types"
        (are [expected cv date-format?] (= expected (read-cell-value cv date-format?))
             2.0 (CellValue. 2.0) false
             "foo" (CellValue. "foo") false
             true (CellValue/valueOf true) false
             date (.. workbook getCreationHelper createFormulaEvaluator (evaluate date-cell)) true))))
             

(deftest read-cell-test
    (let [sheet-data [["Nil" "Blank" "Date" "String" "Number"]
                      [nil "" (july 1) "foo" 42.0]]
          workbook (create-workbook "Sheet 1" sheet-data)
          sheet (.getSheetAt workbook 0)
          rows  (vec (iterator-seq (.iterator sheet)))
          data-row (vec (iterator-seq (.cellIterator (second rows))))
          [nil-cell blank-cell date-cell string-cell number-cell] data-row]
      (testing "Should read all cell types"
        (is (= 2 (count rows)))
        (is (nil? (read-cell nil-cell)))
        (is (= "" (read-cell blank-cell)))
        (is (= (july 1) (read-cell date-cell)))
        (is (= 42.0 (read-cell number-cell))))))

(deftest read-sparse-cell-test
  (let [sheet-data [[nil "temp"]]
        workbook (create-sparse-workbook "Sheet 1" sheet-data)
        sheet (.getSheetAt workbook 0)
        row (.getRow sheet 0)
        nil-cell (.getCell row 0 Row$MissingCellPolicy/RETURN_NULL_AND_BLANK)
        blank-cell (.getCell row 1 Row$MissingCellPolicy/RETURN_NULL_AND_BLANK)
        _ (set-cell! blank-cell nil)
        blank-cell (.getCell row 1 Row$MissingCellPolicy/RETURN_NULL_AND_BLANK)]
    (testing "Should read nil and blank cell types as nil"
      (is (nil? nil-cell))
      (is (nil? (read-cell nil-cell)))
      (is (nil? (read-cell blank-cell))))))

(deftest select-cell-test
  (let [file (config :simple)
        loaded (load-workbook file)
        worksheet (first (sheet-seq loaded))]
    (testing "select-cell"
      (is (= "Empty" (read-cell (select-cell "A1" worksheet))))
      (is (= "Empty" (read-cell (select-cell "B1" worksheet))))
      (is (= 1.0      (read-cell (select-cell "A2" worksheet))))
      (is (= 2.0      (read-cell (select-cell "B2" worksheet))))
      (is (= 3.0      (read-cell (select-cell "B3" worksheet)))))))

(deftest set-cell!-test
  (let [sheet-name "Sheet 1"
        sheet-data [["A1"]]
        workbook (create-workbook sheet-name sheet-data)
        a1 (-> workbook (.getSheetAt 0) (.getRow 0) (.getCell 0))]
    (testing "set-cell! for Date"
      (testing "should set value"
        (set-cell! a1 (july 1))
        (is (= (july 1) (.getDateCellValue a1))))
      (testing "should set nil"
        (let [^java.util.Date nil-date nil]
          (set-cell! a1 nil-date))
        (is (= nil (.getDateCellValue a1)))))
    (testing "set-cell! for String"
      (testing "should set value"
        (set-cell! a1 "foo")
        (is (= "foo" (.getStringCellValue a1)))))
    (testing "set-cell! for boolean"
      (testing "should set value"
        (set-cell! a1 (boolean true))
        (is (.getBooleanCellValue a1))))
    (testing "set-cell! for number"
      (testing "should set int"
        (set-cell! a1 (int 1))
        (is (= 1.0 (.getNumericCellValue a1))))
      (testing "should set double"
        (set-cell! a1 (double 1.2))
        (is (= 1.2 (.getNumericCellValue a1)))))))


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
       (is (= [sheet2 sheet3] (rest actual)))))
    (testing "Should fail on invalid type"
      (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (sheet-seq "not-a-workbook"))))))

(deftest row-seq-test
  (let [sheet-name "Sheet 1"
        sheet-data [["A1" "B1"] ["A2" "B2"]]
        workbook (create-workbook sheet-name sheet-data)
        sheet (select-sheet sheet-name workbook)]
    (testing "Sheet"
      (let [actual (row-seq sheet)]
       (is (= 2 (count actual)))))))

(deftest sparse-row-seq-test
  (let [sheet-name "Sheet 1"
        sheet-data [["A1" "B1"] nil ["A3" "B3"]]
        workbook (create-sparse-workbook sheet-name sheet-data)
        sheet (select-sheet sheet-name workbook)]
    (let [actual (row-seq sheet)]
      (is (= 3 (count actual)))
      (is (empty? (second actual))))))


(deftest cell-seq-test
  (let [sheet-name "Sheet 1"
        sheet-data [["A1" "B1"] ["A2" "B2"]]
        workbook (create-workbook sheet-name sheet-data)
        sheet (select-sheet sheet-name workbook)]
    (testing "for sheet"
      (let [actual (cell-seq sheet)]
       (is (= 4 (count actual)))))
    (testing "for row"
      (let [actual (cell-seq (first (row-seq sheet)))]
       (is (= 2 (count actual)) "Expected correct number of cells.")
       (is (= "A1" (read-cell (first actual))))
       (is (= "B1" (read-cell (second actual))))))
    (testing "for row collection"
      (let [actual (cell-seq (row-seq sheet))]
       (is (= 4 (count actual)) "Expected to get all cells.")
       (is (= ["A1" "B1" "A2" "B2"] (map read-cell actual)))))
    (testing "for sheet collection"
      (let [sheet2 (add-sheet! workbook "Sheet 2")]
       (do (add-rows! sheet2 [["S2/A1" "S2/B1"] ["S2/A2" "S2/B2"]]))
       (let [actual (cell-seq (sheet-seq workbook))]
         (is (= ["A1" "B1"
                 "A2" "B2"
                 "S2/A1" "S2/B1"
                 "S2/A2" "S2/B2"] (map read-cell actual))))))))

(deftest sparse-cell-seq-test
  (let [sheet-name "Sheet 1"
        sheet-data [["A1" nil "C1"] nil [nil "B3" nil "D3"]]
        workbook (create-sparse-workbook sheet-name sheet-data)
        sheet (select-sheet sheet-name workbook)]
    (testing "for sheet"
      (let [actual (cell-seq sheet)]
        (is (= 7 (count actual)))))
    (testing "for row"
      (let [actual (cell-seq (first (row-seq sheet)))
            vals (map #(when % (read-cell %)) actual)]
        (is (= vals ["A1" nil "C1"]))))
    (testing "for row collection"
      (let [actual (cell-seq (row-seq sheet))]
        (is (= 7 (count actual)) "Expected to get all cells.")
        (is (= ["A1" nil "C1" nil "B3" nil "D3"] (map #(when % (read-cell %)) actual)))))))

(deftest sheet-name-test
  (let [name       "Sheet 1"
        data       [["foo" "bar"]]
        workbook   (create-workbook name data)
        sheet      (first (sheet-seq workbook))]
    (is (= name (sheet-name sheet)) "Expected correct sheet name."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (sheet-name "not-a-sheet")))))


(deftest select-sheet-using-string-test
  (let [name       "Sheet 1"
        data       [["foo" "bar"]]
        workbook   (create-workbook name data)
        sheet      (first (sheet-seq workbook))]
    (is (= sheet (select-sheet "Sheet 1" workbook)) "Expected to find the sheet.")
    (is (nil? (select-sheet "unknown name" workbook)) "Expected to get nil for no match."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (select-sheet "name" "not-a-workbook")))))

(deftest select-sheet-using-regex-test
  (let [name       "Sheet 1"
        data       [["foo" "bar"]]
        workbook   (create-workbook name data)
        first-sheet (first (sheet-seq workbook))]
    (is (= first-sheet (select-sheet #"(?i)sheet.*" workbook)) "Expected to find the sheet.")
    (is (nil? (select-sheet #"unknown name" workbook)) "Expected to get nil for no match."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (select-sheet #"name" "not-a-workbook")))))

(deftest select-sheet-using-fn-test
  (let [name       "Sheet 1"
        data       [["foo"] ["bar"]]
        workbook   (create-workbook name data)
        first-sheet (first (sheet-seq workbook))]
    (is (= first-sheet (select-sheet (fn [sheet] (= 2 (count (row-seq sheet)))) workbook)) "Expected to find sheet")
    (is (nil? (select-sheet (constantly false) workbook)) "Expected to get nil for no match."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (select-sheet (constantly true) "not-a-workbook")))))

(deftest select-columns-test
  (let [data     [["Name" "Quantity" "Price" "On Sale"]
                  ["foo" 1.0 42 true]
                  ["bar" 2.0 108 false]]
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
            (second data-rows) (data 2))))
    (testing "Should fail on invalid parameter types."
      (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (select-columns {:A :first, :B :second} "not-a-worksheet"))))))

(deftest sparse-select-columns-test
  (let [data     [["Name" "Quantity" "Thing" "Price" "On Sale"]
                  nil
                  ["foo" 1.0 nil 42 true]
                  ["bar" 2.0 "!" 108 false]]
        workbook (create-workbook "Sheet 1" data)
        sheet    (first (sheet-seq workbook))]
    (testing "Missing rows are skipped in results"
      (let [rows (select-columns {:A :name, :B :quantity} sheet)]
        (is (= [false true false false] (map nil? rows)))
        (are [actual expected] (= actual (zipmap [:name :quantity] expected))
                               (first rows) (data 0)
                               (nth rows 2) (data 2)
                               (nth rows 3) (data 3))))
    (testing "Missing columns should create nil values in map"
      (let [rows (select-columns {:A :name, :C :thing} sheet)]
        (is (= rows [{:name "Name" :thing "Thing"}
                     nil
                     {:name "foo" :thing nil}
                     {:name "bar" :thing "!"}]))))))

(deftest row-seq-test
  (testing "Should fail on invalid parameter types."
    (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (row-seq "not-a-sheet")))))

;; ----------------------------------------------------------------
;; Styling
;; ----------------------------------------------------------------

(deftest create-cell-style!-test
  (testing "Should create a cell style based on the options"
    (testing "no style"
      (let [wb (create-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb)]
       (is (= FillPatternType/NO_FILL (.getFillPatternEnum cs)))
       (is (= false (.getBold (get-font cs wb))))
       (is (= Font/U_NONE (.getUnderline (get-font cs wb))))
        ;Font/COLOR_NORMAL doesn't work right in xssf
       (is (= XSSFFont/DEFAULT_FONT_COLOR (.getColor (get-font cs wb))))
       (is (not (.getItalic (get-font cs wb))))
       (is (not (.getWrapText cs)))
       (is (= HorizontalAlignment/GENERAL (.getAlignmentEnum cs)))
       (is (= BorderStyle/NONE (.getBorderLeftEnum cs)))
       (is (= BorderStyle/NONE (.getBorderRightEnum cs)))
       (is (= BorderStyle/NONE (.getBorderTopEnum cs)))
       (is (= BorderStyle/NONE (.getBorderBottomEnum cs)))
       (is (zero? (.getIndention cs)))
       (is (= "General" (.getDataFormatString cs)))))
    (testing ":background"
      (let [wb (create-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:background :yellow})]
       (is (= FillPatternType/SOLID_FOREGROUND (.getFillPatternEnum cs)))
       (is (= (.getIndex IndexedColors/YELLOW) (.getFillForegroundColor cs)))))
    (testing ":halign"
      (let [wb (create-workbook "Dummy" [["foo"]])
            csl (create-cell-style! wb {:halign :left})
            csr (create-cell-style! wb {:halign :right})
            csc (create-cell-style! wb {:halign :center})]
       (is (= HorizontalAlignment/LEFT (.getAlignmentEnum csl)))
       (is (= HorizontalAlignment/RIGHT (.getAlignmentEnum csr)))
       (is (= HorizontalAlignment/CENTER (.getAlignmentEnum csc)))))
    (testing ":valign"
      (let [wb (create-workbook "Dummy" [["foo"]])
            cst (create-cell-style! wb {:valign :top})
            csb (create-cell-style! wb {:valign :bottom})
            csc (create-cell-style! wb {:valign :center})]
       (is (= VerticalAlignment/TOP (.getVerticalAlignmentEnum cst)))
       (is (= VerticalAlignment/BOTTOM (.getVerticalAlignmentEnum csb)))
       (is (= VerticalAlignment/CENTER (.getVerticalAlignmentEnum csc)))))
    (testing "borders"
      (let [wb (create-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:border-left :thin :border-right :medium
                                       :border-top :thick :border-bottom :thin})]
       (is (= BorderStyle/THIN (.getBorderLeftEnum cs)))
       (is (= BorderStyle/MEDIUM (.getBorderRightEnum cs)))
       (is (= BorderStyle/THICK (.getBorderTopEnum cs)))
       (is (= BorderStyle/THIN (.getBorderBottomEnum cs)))))
    (testing "border colors"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:border-left :thin
                                       :border-right :medium
                                       :border-top :thick 
                                       :border-bottom :thin
                                       :left-border-color :red
                                       :right-border-color :blue
                                       :top-border-color :green
                                       :bottom-border-color :yellow})]
        (is (= BorderStyle/THIN (.getBorderLeftEnum cs)))
        (is (= BorderStyle/MEDIUM (.getBorderRightEnum cs)))
        (is (= BorderStyle/THICK (.getBorderTopEnum cs)))
        (is (= BorderStyle/THIN (.getBorderBottomEnum cs)))
        (is (= (.getIndex IndexedColors/RED) (.getLeftBorderColor cs)))
        (is (= (.getIndex IndexedColors/BLUE) (.getRightBorderColor cs)))
        (is (= (.getIndex IndexedColors/GREEN) (.getTopBorderColor cs)))
        (is (= (.getIndex IndexedColors/YELLOW) (.getBottomBorderColor cs)))))
    (testing ":wrap"
      (let [wb (create-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:wrap :true})]
       (is (.getWrapText cs))))
    (testing ":font :bold"
      (let [wb (create-workbook "Dummy" [["fonts"]])
            fontmap {:bold true}
            testfont (create-font! wb fontmap)
            cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
       (is (.getBold (get-font cs wb)))
       (is (.getBold (get-font cs2 wb)))))
    (testing ":font :color"
      (let [wb (create-workbook "Dummy" [["fonts"]])
            fontmap {:color :light_green}
            testfont (create-font! wb fontmap)
            cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
       (is (= (.getIndex IndexedColors/LIGHT_GREEN) (.getColor (get-font cs wb))))
       (is (= (.getIndex IndexedColors/LIGHT_GREEN) (.getColor (get-font cs2 wb))))))
    (testing ":font :name"
      (let [wb (create-workbook "Dummy" [["fonts"]])
            fontmap {:name "Verdana"}
            testfont (create-font! wb fontmap)
            cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
       (is (= "Verdana" (.getFontName (get-font cs wb))))
       (is (= "Verdana" (.getFontName (get-font cs2 wb))))))
    (testing ":font :size"
      (let [wb (create-workbook "Dummy" [["fonts"]])
            fontmap {:size 8}
            testfont (create-font! wb fontmap)
            cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
       (is (= 8 (.getFontHeightInPoints (get-font cs wb))))
       (is (= 8 (.getFontHeightInPoints (get-font cs2 wb))))))
    (testing ":font :italic"
      (let [wb (create-workbook "Dummy" [["fonts"]])
            fontmap {:italic true}
            testfont (create-font! wb fontmap)
            cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
       (is (.. cs getFont getItalic))
       (is (.. cs2 getFont getItalic))))
    (testing ":font :underline"
      (let [wb (create-workbook "Dummy" [["fonts"]])
            fontmap {:underline true}
            testfont (create-font! wb fontmap)
            cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
       (is (= Font/U_SINGLE (.getUnderline (get-font cs wb))))
       (is (= Font/U_SINGLE (.getUnderline (get-font cs2 wb))))))
    (testing ":indent"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            cs (create-cell-style! wb {:indent 1})
            cs2 (create-cell-style! wb {:indent 2})]
        (is (= 1 (.getIndention cs)))
        (is (= 2 (.getIndention cs2)))))
    (testing ":data-format"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:data-format "#,##0"})
            cs2 (create-cell-style! wb {:data-format "#,##0"})]
        (is (= "#,##0" (.getDataFormatString cs)))
        (is (= (.getDataFormat cs) (.getDataFormat cs2)))))))

(deftest create-font!-test
    (let [wb (create-workbook "Dummy" [["foo"]])]
      (testing "Should create font based on options."
       (let [f-default (create-font! wb {})
             f-not-bold (create-font! wb {:bold false})
             f-bold    (create-font! wb {:bold true})]
         (is (= false (.getBold f-default)))
         (is (= false (.getBold f-not-bold)))
         (is (= true (.getBold f-bold)))))
      (is (thrown-with-msg? IllegalArgumentException #"^workbook.*"
           (create-font! "not-a-workbook" {})))))


(deftest set-cell-style!-test
  (testing "Should apply style to cell."
    (let [wb (create-workbook "Dummy" [["foo"]])
          stylemap {:background :yellow :font {:size 8 :italic true}
                    :wrap true :border-top :medium :valign :center}
          cs (create-cell-style! wb stylemap)
          cell (-> (sheet-seq wb) first cell-seq first)]
      (is (= cell (set-cell-style! cell cs)))
      (is (= (.getCellStyle cell) cs)))))

(deftest set-cell-comment!-test
  (testing "Should set cell comment based on supplied options"
    (testing "comment string"
      (let [wb (create-workbook "Dummy" [["foo00" "foo01" "foo02" "foo03"]
                                         ["foo04" "foo05" "foo06" "foo07"]
                                         ["foo08" "foo09" "foo10" "foo11"]
                                         ["foo12" "foo13" "bar14" "foo15"]
                                         ["foo16" "foo17" "bar18" "foo19"]])
            cellsq (-> (sheet-seq wb) first cell-seq)
            ;cells
            cell00 (nth cellsq 0)
            cell01 (nth cellsq 1)
            cell02 (nth cellsq 2)
            cell03 (nth cellsq 3)
            cell04 (nth cellsq 4)
            cell05 (nth cellsq 5)
            cell06 (nth cellsq 6)
            cell07 (nth cellsq 7)
            cell08 (nth cellsq 8)
            cell09 (nth cellsq 9)
            cell10 (nth cellsq 10)
            cell11 (nth cellsq 11)
            cell12 (nth cellsq 12)
            cell13 (nth cellsq 13)
            cell14 (nth cellsq 14)
            cell15 (nth cellsq 15)
            cell16 (nth cellsq 16)
            comment-str "Short\ncomment."
            empty-str ""
            blank-str " "
            font-bold (create-font! wb {:bold true})
            font-italic (create-font! wb {:italic true})
            font-underline (create-font! wb {:underline true})
            font-name (create-font! wb {:name "Verdana"})
            font-color (create-font! wb {:color :blue})
            font-size (create-font! wb {:size 8})
            ;cell comments
            _ (set-cell-comment! cell00 comment-str)
            _ (set-cell-comment! cell01 empty-str)
            _ (set-cell-comment! cell02 blank-str)
            _ (set-cell-comment! cell03 comment-str :font {:bold true})
            _ (set-cell-comment! cell04 comment-str :font font-bold)
            _ (set-cell-comment! cell05 comment-str :font {:italic true})
            _ (set-cell-comment! cell06 comment-str :font font-italic)
            _ (set-cell-comment! cell07 comment-str :font {:underline true})
            _ (set-cell-comment! cell08 comment-str :font font-underline)
            _ (set-cell-comment! cell09 comment-str :font {:name "Verdana"})
            _ (set-cell-comment! cell10 comment-str :font font-name)
            _ (set-cell-comment! cell11 comment-str :font {:color :blue})
            _ (set-cell-comment! cell12 comment-str :font font-color)
            _ (set-cell-comment! cell13 comment-str :font {:size 8})
            _ (set-cell-comment! cell14 comment-str :font font-size)
            _ (set-cell-comment! cell15 comment-str :width 5)
            _ (set-cell-comment! cell16 comment-str :height 6)
            ;RichTextString instances
            rts00 (.. cell00 getCellComment getString)
            rts01 (.. cell01 getCellComment getString)
            rts02 (.. cell02 getCellComment getString)
            rts03 (.. cell03 getCellComment getString)
            rts04 (.. cell04 getCellComment getString)
            rts05 (.. cell05 getCellComment getString)
            rts06 (.. cell06 getCellComment getString)
            rts07 (.. cell07 getCellComment getString)
            rts08 (.. cell08 getCellComment getString)
            rts09 (.. cell09 getCellComment getString)
            rts10 (.. cell10 getCellComment getString)
            rts11 (.. cell11 getCellComment getString)
            rts12 (.. cell12 getCellComment getString)
            rts13 (.. cell13 getCellComment getString)
            rts14 (.. cell14 getCellComment getString)
            rts15 (.. cell15 getCellComment getString)
            rts16 (.. cell16 getCellComment getString)
            ;extracted cell fonts
            font00 (.getFontAtIndex rts00 0)
            font01 (.getFontAtIndex rts01 0)
            font02 (.getFontAtIndex rts02 0)
            font03 (.getFontAtIndex rts03 0)
            font04 (.getFontAtIndex rts04 0)
            font05 (.getFontAtIndex rts05 0)
            font06 (.getFontAtIndex rts06 0)
            font07 (.getFontAtIndex rts07 0)
            font08 (.getFontAtIndex rts08 0)
            font09 (.getFontAtIndex rts09 0)
            font10 (.getFontAtIndex rts10 0)
            font11 (.getFontAtIndex rts11 0)
            font12 (.getFontAtIndex rts12 0)
            font13 (.getFontAtIndex rts13 0)
            font14 (.getFontAtIndex rts14 0)
            font15 (.getFontAtIndex rts15 0)
            font16 (.getFontAtIndex rts16 0)]
        (is (= comment-str (.getString rts00)))
        (is (= empty-str (.getString rts01)))
        (is (= blank-str (.getString rts02)))
        (is (= true (.getBold font03)))
        (is (= true (.getBold font04)))
        (is (.getItalic font05))
        (is (.getItalic font06))
        (is (= Font/U_SINGLE (.getUnderline font07)))
        (is (= Font/U_SINGLE (.getUnderline font08)))
        (is (= "Verdana" (.getFontName font09)))
        (is (= "Verdana" (.getFontName font10)))
        (is (= (.getIndex IndexedColors/BLUE) (.getColor font11)))
        (is (= (.getIndex IndexedColors/BLUE) (.getColor font12)))
        (is (= 8 (.getFontHeightInPoints font13)))
        (is (= 8 (.getFontHeightInPoints font14)))))))
        ;TODO: test :width and :height options
        

(deftest set-row-style!-test
  (testing "Should apply style to all cells in row."
    (let [wb (create-workbook "Dummy" [["foo" "bar"] ["data b" "data b"]])
          cs (create-cell-style! wb {:background :yellow})
          rs (row-seq (select-sheet "Dummy" wb))
          [header-row, data-row] rs
          [a1, b1] (cell-seq header-row)
          [a2, b2] (cell-seq data-row)]
      (set-row-style! header-row cs)
      (is (= (.getIndex IndexedColors/YELLOW) (.. a1 getCellStyle getFillForegroundColor)))
      (is (= (.getIndex IndexedColors/YELLOW) (.. b1 getCellStyle getFillForegroundColor)))
      (is (not= (.getIndex IndexedColors/YELLOW) (.. a2 getCellStyle getFillForegroundColor)))
      (is (not= (.getIndex IndexedColors/YELLOW) (.. b2 getCellStyle getFillForegroundColor))))))
      

(deftest set-row-styles!-test
  (testing "Should apply the given styles to the row's cells in order."
    (let [wb (create-workbook "Dummy" [["foo" "bar"] ["data b" "data b"]])
          cs1 (create-cell-style! wb {:background :yellow})
          cs2 (create-cell-style! wb {:background :red})
          rs (row-seq (select-sheet "Dummy" wb))
          [header-row, data-row] rs
          [a1, b1] (cell-seq header-row)
          [a2, b2] (cell-seq data-row)]
      (do (set-row-styles! header-row (list cs1 cs2)))
      (is (= (.getIndex IndexedColors/YELLOW) (.. a1 getCellStyle getFillForegroundColor)))
      (is (= (.getIndex IndexedColors/RED) (.. b1 getCellStyle getFillForegroundColor)))
      (is (not= (.getIndex IndexedColors/YELLOW) (.. a2 getCellStyle getFillForegroundColor)))
      (is (not= (.getIndex IndexedColors/RED) (.. b2 getCellStyle getFillForegroundColor))))))
      

(deftest get-row-styles-test
  (testing "Should get a seq of the row's CellStyles."
    (let [wb (create-workbook "Dummy" [["foo" "bar"] ["data b" "data b"]])
          cs1 (create-cell-style! wb {:background :yellow})
          cs2 (create-cell-style! wb {:background :red})
          rs (row-seq (select-sheet "Dummy" wb))
          [header-row, data-row] rs
          [a1, b1] (cell-seq header-row)
          [a2, b2] (cell-seq data-row)]
      (do (set-row-styles! header-row (list cs1 cs2)))
      (is (= (list cs1 cs2) (get-row-styles header-row)))
      (is (= (list cs1 cs2) (get-row-styles header-row))))))
      

(deftest cell-fn-test
  (testing "Creating a function from a formula cell"
    (let [file (config :simple)
           loaded (load-workbook file)
           worksheet (first (sheet-seq loaded))
           cell-function (cell-fn "B3" worksheet "A2")]
       (is (= (cell-function 2.0) 5.0))
       (is (= (cell-function 3.0) 7.0)))))
       

;; ----------------------------------------------------------------
;; Integration tests
;; ----------------------------------------------------------------

(defn- test-loaded-workbook [loaded]
  (is (isa? (class loaded) Workbook)))

(deftest load-workbook-from-stream-integration-test
  (with-open [stream (FileInputStream. (config :datatypes-file))]
    (let [loaded (load-workbook-from-stream stream)]
      (test-loaded-workbook loaded))))

(deftest load-workbook-from-file-integration-test
  (let [file (config :datatypes-file)
        loaded (load-workbook-from-file file)]
    (test-loaded-workbook loaded)))

(defn- path->dir-and-file
  [^String path]
  (let [i (.lastIndexOf path "/")
        dir (.substring path 0 i)
        file (.substring path (inc i))]
    [dir file]))

(deftest load-workbook-integration-test
  (testing "should accept file name as string"
    (let [file (config :datatypes-file)
          loaded (load-workbook file)]
      (test-loaded-workbook loaded)))
  (testing "should accept InputStream"
    (with-open [stream (FileInputStream. (config :datatypes-file))]
      (let [loaded (load-workbook stream)]
        (test-loaded-workbook loaded)))))

(deftest load-workbook-from-resource-integration-test
  (let [file (config :datatypes-test-resource)
        loaded (load-workbook-from-resource file)]
    (test-loaded-workbook loaded)))

(deftest save-workbook!-test
  (testing "Should fail on invalid parameter types."
    (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (save-workbook! "filename.xlsx" "not-a-workbook"))))
  (testing "Saving workbook into a file"
    (let [file (config :save-workbook-location)
          workbook (create-workbook "Sheet 1" [["A1" "B1" "C1"]])
          _ (save-workbook! file workbook)
          loaded (load-workbook file)
          _ (io/delete-file file)]
      (test-loaded-workbook loaded)))
  (testing "Saving workbook into a stream"
    (let [file (config :save-workbook-location)
          stream (io/output-stream file)
          workbook (create-workbook "Sheet 1" [["A1" "B1" "C1"]])
          _ (save-workbook! stream workbook)
          loaded (load-workbook file)
          _ (io/delete-file file)]
      (test-loaded-workbook loaded))))

(defn- datatypes-rows [file]
  (->> (load-workbook-from-file file)
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
      (is (every? number? (datatypes-data file :scientific)))
      (is (every? date? (datatypes-data file :date-formulae))))))

(deftest select-columns-formula-evaluation-integration-test
  (testing "Formula evaluation"
    (let [file (config :formulae-file)
          formula-expected-pairs (->> (load-workbook file)
                                  sheet-seq
                                  first
                                  (select-columns formulae-map)
                                  rest)]
      (is (every? #(= (:formula %) (:expected %)) formula-expected-pairs)))))

(deftest select-blanks-integration-test
  (let [file (config :blanks-file)
        workbook (load-workbook file)
        sheet (select-sheet "test_data" workbook)]
    (testing "reading rows"
      (let [actual (row-seq sheet)]
        (is (= 5 (count actual)))
        (is (nil? (second actual)))))
    (testing "reading cells"
      (let [row (.getRow sheet 2)
            actual (cell-seq row)]
        (is (= 5 (count actual)))
        (is (= ["mid left" nil "middle" nil "mid right"] (map #(when % (read-cell %)) actual)))))
    (testing "selecting columns"
      (is (= [{:A "top left"}
              {:E "mid right", :C "middle", :A "mid left"}
              {:E "bottom right"}]
             (select-columns {:A :A :B :B :C :C :D :D :E :E} sheet))))))

(deftest name-test
  (let [data [["Test1"  "First"    "Second"]
              ["Test2"  "Third"    "Fourth"]
              [nil      "Fifth"    "Sixth"]
              [nil      "Seventh"  "Eight"]
              [nil      "Ninth"    "Tenth"]]
        workbook (create-workbook "Sheet 1" data)]
    (testing "Set named range and retrieve cells from it."
             (add-name! workbook "test1" "'Sheet 1'!$A$1")
             (add-name! workbook "ten" "'Sheet 1'!$B$1:$C$5")
             (is (= "Test1" (read-cell (first (select-name workbook "test1")))))
             (is (= (reduce concat (map (fn [[_ a b]] [a b]) data))
                    (map read-cell (select-name workbook "ten"))))
             (is (nil? (select-name workbook "bill"))))))

(deftest date-bases-test
  (letfn [(read-sheet [file]
            (->> (load-workbook file)
                 sheet-seq
                 first
                 (select-columns {:A :date, :B :year, :C :comment})
                 rest))
          (year [^java.util.Date date]
            (+ 1900 (.getYear date)))]
    (testing "Can read workbooks with 1900-based dates"
      (let [actual (read-sheet (config :1900-based-file))]
        (is (every? #(== (year (:date %)) (:year %)) actual))))
    (testing "Can read workbooks with 1904-based dates"
      (let [actual (read-sheet (config :1904-based-file))]
        (is (every? #(== (year (:date %)) (:year %)) actual))))))

(deftest select-cell-update-value-read-updated-formula-test
  (let [file (config :simple)
        loaded (load-workbook file)
        worksheet (first (sheet-seq loaded))]
    (testing "selecting-cell"
      (is (= 1.0      (read-cell (select-cell "A2" worksheet))))
      (testing "updating cell-value"
        (set-cell! (select-cell "A2" worksheet) 2.0)
        (is (= 2.0      (read-cell (select-cell "A2" worksheet))))
        (is (= 3.0      (read-cell (select-cell "B2" worksheet))))
        (is (= 5.0      (read-cell (select-cell "B3" worksheet))))))))

(deftest select-cell-overwrite-formula-read-updated-formula-test
  (let [file (config :simple)
        loaded (load-workbook file)
        worksheet (first (sheet-seq loaded))]
    (testing "selecting-cell"
      (is (= 2.0      (read-cell (select-cell "B2" worksheet))))
      (testing "updating cell-value"
        (set-cell! (select-cell "B2" worksheet) 4.0)
        (is (= 1.0      (read-cell (select-cell "A2" worksheet))))
        (is (= 4.0      (read-cell (select-cell "B2" worksheet))))
        (is (= 5.0      (read-cell (select-cell "B3" worksheet))))))))

(deftest no-style-overflows-test
  (testing "Can add many date cells without overflowing"
    (let [file (config :save-workbook-location)
          workbook (create-workbook "Sheet 1"
                                    (mapv (fn [i] [i #inst "2018-10-01T00:00:00Z"])
                                          (range 100000)))
          _ (save-workbook! file workbook)
          loaded (load-workbook file)
          _ (io/delete-file file)]
      (test-loaded-workbook loaded))))

(ns dk.ative.docjure.xls-test
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:use [clojure.test])
  (:import (org.apache.poi.ss.usermodel Workbook Sheet Cell Row CellStyle IndexedColors Font CellValue)
	   (org.apache.poi.hssf.usermodel HSSFWorkbook)
	   (java.util Date)))

(def config {:datatypes-file "test/dk/ative/docjure/testdata/datatypes.xls"
	     :formulae-file "test/dk/ative/docjure/testdata/formulae.xls"})
(def datatypes-map {:A :text, :B :integer, :C :decimal, :D :date, :E :time, :F :date-time, :G :percentage, :H :fraction, :I :scientific})
(def formulae-map {:A :formula, :B :expected})

(deftest add-sheet!-test
  (let [workbook (HSSFWorkbook.)
	sheet-name "Sheet 1"
	actual   (add-sheet! workbook sheet-name)]
    (testing "Sheet creation"
      (is (= 1 (.getNumberOfSheets workbook)) "Expected sheet to be added.")
      (is (= sheet-name (.. workbook (getSheetAt 0) (getSheetName))) "Expected sheet to have correct name."))
    (testing "Should fail on non-Workbook"
      (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (add-sheet! "not-a-workbook" "sheet-name"))))))

(deftest create-xls-workbook-test
  (let [sheet-name "Sheet 1" 
	sheet-data [["A1" "B1" "C1"]
		    ["A2" "B2" "C2"]]
	workbook (create-xls-workbook sheet-name sheet-data)]
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

(deftest remove-row!-test
  (let [sheet-name "Sheet 1" 
	sheet-data [["A1" "B1" "C1"]
		    ["A2" "B2" "C2"]]
	workbook (create-xls-workbook sheet-name sheet-data)
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
	workbook (create-xls-workbook sheet-name sheet-data)
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
          workbook (create-xls-workbook "Just a date" [[date]])
          sheet (.getSheetAt workbook 0)
          rows  (vec (iterator-seq (.iterator sheet)))
          data-row (vec (iterator-seq (.cellIterator (first rows))))
          date-cell (first data-row)]
      (testing "Should read all cell types"
        (are [expected cv date-format?] (= expected (read-cell-value cv date-format?))
             2.0 (CellValue. 2.0) false
             "foo" (CellValue. "foo") false
             true (CellValue/valueOf true) false
             date (.. workbook getCreationHelper createFormulaEvaluator (evaluate date-cell)) true
             ))))

(deftest read-cell-test
    (let [sheet-data [["Nil" "Blank" "Date" "String" "Number"]
                      [nil "" (july 1) "foo" 42.0]]
          workbook (create-xls-workbook "Sheet 1" sheet-data)
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

(deftest set-cell!-test
  (let [sheet-name "Sheet 1" 
	sheet-data [["A1"]]
	workbook (create-xls-workbook sheet-name sheet-data)
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
      (let [workbook (HSSFWorkbook.)
	    actual (sheet-seq workbook)]
	(is (not (nil? actual)))
	(is (empty? actual))))
    (testing "Single sheet."
      (let [workbook (create-xls-workbook sheet-name sheet-data)
	    actual   (sheet-seq workbook)]
	(is (= 1 (count actual)))
	(is (= sheet-name (.getSheetName (first actual))))))
    (testing "Multiple sheets."
      (let [workbook (create-xls-workbook sheet-name sheet-data)
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
	workbook (create-xls-workbook sheet-name sheet-data)
	sheet (select-sheet sheet-name workbook)]
    (testing "Sheet"
      (let [actual (row-seq sheet)]
	(is (= 2 (count actual)))))))

(deftest cell-seq-test
  (let [sheet-name "Sheet 1" 
	sheet-data [["A1" "B1"] ["A2" "B2"]]
	workbook (create-xls-workbook sheet-name sheet-data)
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


(deftest sheet-name-test
  (let [name       "Sheet 1" 
	data       [["foo" "bar"]]
	workbook   (create-xls-workbook name data)
	sheet      (first (sheet-seq workbook))]
    (is (= name (sheet-name sheet)) "Expected correct sheet name."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"sheet.*" (sheet-name "not-a-sheet")))))

(deftest select-sheet-using-string-test
  (let [name       "Sheet 1"
	data       [["foo" "bar"]]
	workbook   (create-xls-workbook name data)
	sheet      (first (sheet-seq workbook))]
    (is (= sheet (select-sheet "Sheet 1" workbook)) "Expected to find the sheet.")
    (is (nil? (select-sheet "unknown name" workbook)) "Expected to get nil for no match."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (select-sheet "name" "not-a-workbook")))))

(deftest select-sheet-using-regex-test
  (let [name       "Sheet 1"
	data       [["foo" "bar"]]
	workbook   (create-xls-workbook name data)
	first-sheet (first (sheet-seq workbook))]
    (is (= first-sheet (select-sheet #"(?i)sheet.*" workbook)) "Expected to find the sheet.")
    (is (nil? (select-sheet #"unknown name" workbook)) "Expected to get nil for no match."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (select-sheet #"name" "not-a-workbook")))))

(deftest select-sheet-using-fn-test
  (let [name       "Sheet 1"
	data       [["foo"] ["bar"]]
	workbook   (create-xls-workbook name data)
	first-sheet (first (sheet-seq workbook))]
    (is (= first-sheet (select-sheet (fn [sheet] (= 2 (count (row-seq sheet)))) workbook)) "Expected to find sheet")
    (is (nil? (select-sheet (constantly false) workbook)) "Expected to get nil for no match."))
  (testing "Should fail on invalid parameter type"
    (is (thrown-with-msg? IllegalArgumentException #"workbook.*" (select-sheet (constantly true) "not-a-workbook")))))

(deftest select-columns-test
  (let [data     [["Name" "Quantity" "Price" "On Sale"] 
		  ["foo" 1.0 42 true] 
		  ["bar" 2.0 108 false]]
	workbook (create-xls-workbook "Sheet 1" data)
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

;; ----------------------------------------------------------------
;; Styling
;; ----------------------------------------------------------------

(deftest create-cell-style!-test
  (testing "Should create a cell style based on the options"
    (testing "no style"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb)]
	(is (= CellStyle/NO_FILL (.getFillPattern cs)))
        (is (= Font/BOLDWEIGHT_NORMAL (.getBoldweight (.getFont cs wb))))
        (is (= Font/U_NONE (.getUnderline (.getFont cs wb))))
        (is (= Font/COLOR_NORMAL (.getColor (.getFont cs wb))))
        (is (not (.getItalic (.getFont cs wb))))
        (is (not (.getWrapText cs)))
        (is (= CellStyle/ALIGN_GENERAL (.getAlignment cs)))
        (is (= CellStyle/BORDER_NONE (.getBorderLeft cs)))
        (is (= CellStyle/BORDER_NONE (.getBorderRight cs)))
        (is (= CellStyle/BORDER_NONE (.getBorderTop cs)))
        (is (= CellStyle/BORDER_NONE (.getBorderBottom cs)))))
    (testing ":background"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:background :yellow})]
	(is (= CellStyle/SOLID_FOREGROUND (.getFillPattern cs)))
        (is (= (.getIndex IndexedColors/YELLOW) (.getFillForegroundColor cs)))))
    (testing ":halign"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            csl (create-cell-style! wb {:halign :left})
            csr (create-cell-style! wb {:halign :right})
            csc (create-cell-style! wb {:halign :center})]
	(is (= CellStyle/ALIGN_LEFT (.getAlignment csl)))
        (is (= CellStyle/ALIGN_RIGHT (.getAlignment csr)))
        (is (= CellStyle/ALIGN_CENTER (.getAlignment csc)))))
    (testing ":valign"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cst (create-cell-style! wb {:valign :top})
            csb (create-cell-style! wb {:valign :bottom})
            csc (create-cell-style! wb {:valign :center})]
	(is (= CellStyle/VERTICAL_TOP (.getVerticalAlignment cst)))
        (is (= CellStyle/VERTICAL_BOTTOM (.getVerticalAlignment csb)))
        (is (= CellStyle/VERTICAL_CENTER (.getVerticalAlignment csc)))))
    (testing "borders"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:border-left :thin :border-right :medium
                                       :border-top :thick :border-bottom :thin})]
	(is (= CellStyle/BORDER_THIN (.getBorderLeft cs)))
        (is (= CellStyle/BORDER_MEDIUM (.getBorderRight cs)))
        (is (= CellStyle/BORDER_THICK (.getBorderTop cs)))
        (is (= CellStyle/BORDER_THIN (.getBorderBottom cs)))))
    (testing ":wrap"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:wrap :true})]
        (is (.getWrapText cs))))
    (testing ":font :bold"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:bold true}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= Font/BOLDWEIGHT_BOLD (.getBoldweight (.getFont cs wb))))
        (is (= Font/BOLDWEIGHT_BOLD (.getBoldweight (.getFont cs2 wb))))))
    (testing ":font :color"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:color :light_green}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= (.getIndex IndexedColors/LIGHT_GREEN) (.getColor (.getFont cs wb))))
        (is (= (.getIndex IndexedColors/LIGHT_GREEN) (.getColor (.getFont cs2 wb))))))
    (testing ":font :name"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:name "Verdana"}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= "Verdana" (.getFontName (.getFont cs wb))))
        (is (= "Verdana" (.getFontName (.getFont cs2 wb))))))
    (testing ":font :size"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:size 8}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= 8 (.getFontHeightInPoints (.getFont cs wb))))
        (is (= 8 (.getFontHeightInPoints (.getFont cs2 wb))))))
    (testing ":font :italic"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:italic true}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (.getItalic (.getFont cs wb)))
        (is (.getItalic (.getFont cs2 wb)))))
    (testing ":font :underline"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:underline true}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= Font/U_SINGLE (.getUnderline (.getFont cs wb))))
        (is (= Font/U_SINGLE (.getUnderline (.getFont cs2 wb))))))))

(deftest create-font!-test
    (let [wb (create-xls-workbook "Dummy" [["foo"]])]
      (testing "Should create font based on options."
	(let [f-default (create-font! wb {})
	      f-not-bold (create-font! wb {:bold false})
	      f-bold    (create-font! wb {:bold true})]
	  (is (= Font/BOLDWEIGHT_NORMAL (.getBoldweight f-default)))
	  (is (= Font/BOLDWEIGHT_NORMAL (.getBoldweight f-not-bold)))
	  (is (= Font/BOLDWEIGHT_BOLD (.getBoldweight f-bold)))))
      (is (thrown-with-msg? IllegalArgumentException #"^workbook.*"
	    (create-font! "not-a-workbook" {})))))
	    

(deftest set-cell-style!-test
  (testing "Should apply style to cell."
    (let [wb (create-xls-workbook "Dummy" [["foo"]])
          stylemap {:background :yellow :font {:size 8 :italic true}
                    :wrap true :border-top :medium :valign :center}
	  cs (create-cell-style! wb stylemap)
	  cell (-> (sheet-seq wb) first cell-seq first)]
      (is (= cell (set-cell-style! cell cs)))
      (is (= (.getCellStyle cell) cs)))))

(deftest set-cell-comment!-test
  (testing "Should set cell comment."
    (let [wb (create-xls-workbook "Dummy" [["foo"]])
	  cell (-> (sheet-seq wb) first cell-seq first)
          comment-str "This is a\nshort comment."
          cellc (set-cell-comment! cell comment-str)
          cstr (.. cell getCellComment getString getString)]
      (is (= comment-str cstr)))))

(deftest set-row-style!-test
  (testing "Should apply style to all cells in row."
    (let [wb (create-xls-workbook "Dummy" [["foo" "bar"] ["data b" "data b"]])
	  cs (create-cell-style! wb {:background :yellow})
	  rs (row-seq (select-sheet "Dummy" wb))
	  [header-row, data-row] rs
	  [a1, b1] (cell-seq header-row)
	  [a2, b2] (cell-seq data-row)]
      (do (set-row-style! header-row cs))
      (is (= (.getIndex IndexedColors/YELLOW) (.. a1 getCellStyle getFillForegroundColor)))
      (is (= (.getIndex IndexedColors/YELLOW) (.. b1 getCellStyle getFillForegroundColor)))
      (is (not= (.getIndex IndexedColors/YELLOW) (.. a2 getCellStyle getFillForegroundColor)))
      (is (not= (.getIndex IndexedColors/YELLOW) (.. b2 getCellStyle getFillForegroundColor)))
      )))

(deftest set-row-styles!-test
  (testing "Should apply the given styles to the row's cells in order."
    (let [wb (create-xls-workbook "Dummy" [["foo" "bar"] ["data b" "data b"]])
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
      (is (not= (.getIndex IndexedColors/RED) (.. b2 getCellStyle getFillForegroundColor)))
      )))

(deftest get-row-styles-test
  (testing "Should get a seq of the row's CellStyles."
    (let [wb (create-xls-workbook "Dummy" [["foo" "bar"] ["data b" "data b"]])
	  cs1 (create-cell-style! wb {:background :yellow})
	  cs2 (create-cell-style! wb {:background :red})
	  rs (row-seq (select-sheet "Dummy" wb))
	  [header-row, data-row] rs
	  [a1, b1] (cell-seq header-row)
	  [a2, b2] (cell-seq data-row)]
      (do (set-row-styles! header-row (list cs1 cs2)))
      (is (= (list cs1 cs2) (get-row-styles header-row)))
      (is (= (list cs1 cs2) (get-row-styles header-row)))
      )))

;; ----------------------------------------------------------------
;; Integration tests
;; ----------------------------------------------------------------

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

(deftest select-columns-formula-evaluation-integration-test
  (testing "Formula evaluation"
    (let [file (config :formulae-file)
	  formula-expected-pairs (->> (load-workbook file)
				      sheet-seq
				      first
				      (select-columns formulae-map)
				      rest)]
      (is (every? #(= (:formula %) (:expected %)) formula-expected-pairs)))))

(deftest name-test
  (let [data [["Test1"  "First"    "Second"]
              ["Test2"  "Third"    "Fourth"]
              [nil      "Fifth"    "Sixth"]
              [nil      "Seventh"  "Eight"]
              [nil      "Ninth"    "Tenth"]]
	workbook (create-xls-workbook "Sheet 1" data)]
    (testing "Set named range and retrieve cells from it."
             (add-name! workbook "test1" "'Sheet 1'!$A$1")
             (add-name! workbook "ten" "'Sheet 1'!$B$1:$C$5")
             (is (= "Test1" (read-cell (first (select-name workbook "test1")))))
             (is (= (reduce concat (map (fn [[_ a b]] [a b]) data))
                    (map read-cell (select-name workbook "ten"))))
             (is (nil? (select-name workbook "bill"))))))


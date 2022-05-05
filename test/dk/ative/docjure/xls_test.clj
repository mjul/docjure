(ns dk.ative.docjure.xls-test
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:use [clojure.test])
  (:import (org.apache.poi.ss.usermodel Workbook Sheet Cell Row CellStyle IndexedColors
                                        Font CellValue FillPatternType
                                        HorizontalAlignment VerticalAlignment BorderStyle)
	   (org.apache.poi.hssf.usermodel HSSFWorkbook HSSFFont)
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
	(is (= FillPatternType/NO_FILL (.getFillPattern cs)))
        (is (= false (.getBold (get-font cs wb))))
        (is (= Font/U_NONE (.getUnderline (get-font cs wb))))
        (is (= Font/COLOR_NORMAL (.getColor (get-font cs wb))))
        (is (not (.getItalic (get-font cs wb))))
        (is (not (.getWrapText cs)))
        (is (= HorizontalAlignment/GENERAL (.getAlignment cs)))
        (is (= BorderStyle/NONE (.getBorderLeft cs)))
        (is (= BorderStyle/NONE (.getBorderRight cs)))
        (is (= BorderStyle/NONE (.getBorderTop cs)))
        (is (= BorderStyle/NONE (.getBorderBottom cs)))
        (is (zero? (.getIndention cs)))
        (is (= "General" (.getDataFormatString cs)))))
    (testing ":background"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:background :yellow})]
	(is (= FillPatternType/SOLID_FOREGROUND (.getFillPattern cs)))
        (is (= (.getIndex IndexedColors/YELLOW) (.getFillForegroundColor cs)))))
    (testing ":halign"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            csl (create-cell-style! wb {:halign :left})
            csr (create-cell-style! wb {:halign :right})
            csc (create-cell-style! wb {:halign :center})]
	(is (= HorizontalAlignment/LEFT (.getAlignment csl)))
        (is (= HorizontalAlignment/RIGHT (.getAlignment csr)))
        (is (= HorizontalAlignment/CENTER (.getAlignment csc)))))
    (testing ":valign"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cst (create-cell-style! wb {:valign :top})
            csb (create-cell-style! wb {:valign :bottom})
            csc (create-cell-style! wb {:valign :center})]
	(is (= VerticalAlignment/TOP (.getVerticalAlignment cst)))
        (is (= VerticalAlignment/BOTTOM (.getVerticalAlignment csb)))
        (is (= VerticalAlignment/CENTER (.getVerticalAlignment csc)))))
    (testing "borders"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
            cs (create-cell-style! wb {:border-left :thin :border-right :medium
                                       :border-top :thick :border-bottom :thin})]
	(is (= BorderStyle/THIN (.getBorderLeft cs)))
        (is (= BorderStyle/MEDIUM (.getBorderRight cs)))
        (is (= BorderStyle/THICK (.getBorderTop cs)))
        (is (= BorderStyle/THIN (.getBorderBottom cs)))))
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
        (is (= BorderStyle/THIN (.getBorderLeft cs)))
        (is (= BorderStyle/MEDIUM (.getBorderRight cs)))
        (is (= BorderStyle/THICK (.getBorderTop cs)))
        (is (= BorderStyle/THIN (.getBorderBottom cs)))
        (is (= (.getIndex IndexedColors/RED) (.getLeftBorderColor cs)))
        (is (= (.getIndex IndexedColors/BLUE) (.getRightBorderColor cs)))
        (is (= (.getIndex IndexedColors/GREEN) (.getTopBorderColor cs)))
        (is (= (.getIndex IndexedColors/YELLOW) (.getBottomBorderColor cs)))))
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
	(is (= true (.getBold (get-font cs wb))))
        (is (= true (.getBold (get-font cs2 wb))))))
    (testing ":font :color"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:color :light_green}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= (.getIndex IndexedColors/LIGHT_GREEN) (.getColor (get-font cs wb))))
        (is (= (.getIndex IndexedColors/LIGHT_GREEN) (.getColor (get-font cs2 wb))))))
    (testing ":font :name"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:name "Verdana"}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= "Verdana" (.getFontName (get-font cs wb))))
        (is (= "Verdana" (.getFontName (get-font cs2 wb))))))
    (testing ":font :size"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:size 8}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= 8 (.getFontHeightInPoints (get-font cs wb))))
        (is (= 8 (.getFontHeightInPoints (get-font cs2 wb))))))
    (testing ":font :italic"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:italic true}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (.getItalic (get-font cs wb)))
        (is (.getItalic (get-font cs2 wb)))))
    (testing ":font :underline"
      (let [wb (create-xls-workbook "Dummy" [["fonts"]])
            fontmap {:underline true}
            testfont (create-font! wb fontmap)
	    cs (create-cell-style! wb {:font testfont})
            cs2 (create-cell-style! wb {:font fontmap})]
	(is (= Font/U_SINGLE (.getUnderline (get-font cs wb))))
        (is (= Font/U_SINGLE (.getUnderline (get-font cs2 wb))))))
    (testing ":indent"
      (let [wb (create-xls-workbook "Dummy" [["foo"]])
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
    (let [wb (create-xls-workbook "Dummy" [["foo"]])]
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
    (let [wb (create-xls-workbook "Dummy" [["foo"]])
          stylemap {:background :yellow :font {:size 8 :italic true}
                    :wrap true :border-top :medium :valign :center}
	  cs (create-cell-style! wb stylemap)
	  cell (-> (sheet-seq wb) first cell-seq first)]
      (is (= cell (set-cell-style! cell cs)))
      (is (= (.getCellStyle cell) cs)))))

(deftest set-cell-comment!-test
  (testing "Should set cell comment based on supplied options"
    (testing "comment string"
      (let [wb (create-xls-workbook "Dummy"
                                    [["foo00" "foo01" "foo02" "foo03"]
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
            font-bold-idx (.getIndex font-bold)
            font-italic (create-font! wb {:italic true})
            font-italic-idx (.getIndex font-italic)
            font-underline (create-font! wb {:underline true})
            font-underline-idx (.getIndex font-underline)
            font-name (create-font! wb {:name "Verdana"})
            font-name-idx (.getIndex font-name)
            font-color (create-font! wb {:color :blue})
            font-color-idx (.getIndex font-color)
            font-size (create-font! wb {:size 8})
            font-size-idx (.getIndex font-size)
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
            ;;extracted cell fonts
            font00 (.getFontAt wb (.intValue (.getFontAtIndex rts00 0)))
            font01 (.getFontAt wb (.intValue (.getFontAtIndex rts01 0)))
            font02 (.getFontAt wb (.intValue (.getFontAtIndex rts02 0)))
            font03 (.getFontAt wb (.intValue (.getFontAtIndex rts03 0)))
            font04 (.getFontAt wb (.intValue (.getFontAtIndex rts04 0)))
            font05 (.getFontAt wb (.intValue (.getFontAtIndex rts05 0)))
            font06 (.getFontAt wb (.intValue (.getFontAtIndex rts06 0)))
            font07 (.getFontAt wb (.intValue (.getFontAtIndex rts07 0)))
            font08 (.getFontAt wb (.intValue (.getFontAtIndex rts08 0)))
            font09 (.getFontAt wb (.intValue (.getFontAtIndex rts09 0)))
            font10 (.getFontAt wb (.intValue (.getFontAtIndex rts10 0)))
            font11 (.getFontAt wb (.intValue (.getFontAtIndex rts11 0)))
            font12 (.getFontAt wb (.intValue (.getFontAtIndex rts12 0)))
            font13 (.getFontAt wb (.intValue (.getFontAtIndex rts13 0)))
            font14 (.getFontAt wb (.intValue (.getFontAtIndex rts14 0)))
            font15 (.getFontAt wb (.intValue (.getFontAtIndex rts15 0)))
            font16 (.getFontAt wb (.intValue (.getFontAtIndex rts16 0)))]
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
        (is (= 8 (.getFontHeightInPoints font14)))
        ;TODO: test :width and :height options
        ))))

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

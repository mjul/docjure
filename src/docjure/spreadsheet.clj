(ns docjure.spreadsheet
  (:import
    (java.io FileOutputStream FileInputStream)
    (java.util Date Calendar)
    (org.apache.poi.xssf.usermodel XSSFWorkbook)
    (org.apache.poi.ss.usermodel Workbook Sheet Cell Row WorkbookFactory DateUtil)
    (org.apache.poi.ss.util CellReference)))

(defn cell-reference [cell]
  (.formatAsString (CellReference. (.getRowIndex cell) (.getColumnIndex cell))))

(defmulti read-cell #(.getCellType %))
(defmethod read-cell Cell/CELL_TYPE_BLANK     [_]     "")
(defmethod read-cell Cell/CELL_TYPE_STRING    [cell]  (.getStringCellValue cell))
(defmethod read-cell Cell/CELL_TYPE_FORMULA   [cell]  (.getCellFormula cell))
(defmethod read-cell Cell/CELL_TYPE_BOOLEAN   [cell]  (.getBooleanCellValue cell))
(defmethod read-cell Cell/CELL_TYPE_NUMERIC   [cell]
  (if (DateUtil/isCellDateFormatted cell)
    (.getDateCellValue cell)
    (.getNumericCellValue cell)))

(defn load-workbook 
  "Load an Excel .xls or .xlsx workbook from a file."
  [filename]
  (with-open [stream (FileInputStream. filename)]
    (WorkbookFactory/create stream)))

(defn save-workbook! 
  "Save the workbook into a file."
  [filename #^Workbook workbook]
  (with-open [file-out (FileOutputStream. filename)]
    (.write workbook file-out)))

(defn sheet-seq 
  "Return a seq of the sheets in a workbook."
  [#^Workbook workbook]
  (for [idx (range (.getNumberOfSheets workbook))]
    (.getSheetAt workbook idx)))

(defn sheet-name
  "Return the name of a worksheet."
  [#^Sheet sheet]
  (.getSheetName sheet))

(defn select-sheet 
  "Select a sheet from the workbook by name."
  [name workbook]
  (->> (sheet-seq workbook)
       (filter #(= name (sheet-name %)))
       first))

(defn row-seq 
  "Return a sequence of the rows in a sheet."
  [#^Sheet sheet]
  (iterator-seq (.iterator sheet)))

(defn cell-seq
  "Return a seq of the cells in one or more sheets, ordered by row and column."
  [#^Sheet sheet-or-coll]
  (for [sheet (if (seq? sheet-or-coll) sheet-or-coll (list sheet-or-coll))
	row   (iterator-seq (.iterator sheet))
	cell   (iterator-seq (.iterator row))]
    cell))

(defn into-seq
  [sheet-or-row]
  (vec (for [item (iterator-seq (.iterator sheet-or-row))] item)))

(defn- project-cell [column-map cell]
   (let [colname (-> cell
		     .getColumnIndex 
		     org.apache.poi.ss.util.CellReference/convertNumToColString
		     keyword)
	 new-key (column-map colname)]
	 (when new-key
	   {new-key (read-cell cell)})))


(defn select-columns [column-map sheet]
  "Takes two arguments: column hashmap where the keys are the
   spreadsheet column names as keys and the values represent the names they are mapped to, 
   and a sheet.

   For example, to select columns A and C as :first and :third from the sheet
   
   (select-columns {:A :first, :C :third} sheet)
   => [{:first \"Value in cell A1\", :third \"Value in cell C1\"} ...] "
  (vec
   (for [row (into-seq sheet)]
     (->> (map #(project-cell column-map %) row)
	  (apply merge)))))

(defn string-cell? [cell]
  (= Cell/CELL_TYPE_STRING (.getCellType cell)))

(defn- date-or-calendar? [value]
  (let [cls (class value)]
    (or (isa? cls Date) (isa? cls Calendar))))

(defn apply-date-format! [cell format]
  (let [workbook (.. cell getSheet getWorkbook)
        date-style (.createCellStyle workbook)
        format-helper (.getCreationHelper workbook)]
    (.setDataFormat date-style
		    (.. format-helper createDataFormat (getFormat format)))
    (.setCellStyle cell date-style)))

(defn set-cell! [cell value]
  (let [converted-value (cond (number? value) (double value)
                          true value)]
    (.setCellValue cell converted-value)
    (if (date-or-calendar? value)
      (apply-date-format! cell "m/d/yy"))))

(defn add-row! [sheet values]
  (let [row (.createRow sheet (inc (.getLastRowNum sheet)))]
    (doseq [[column-index value] (partition 2 (interleave (iterate inc 0) values))]
      (set-cell! (.createCell row column-index) value))
    row))

(defn add-rows! [sheet rows]
  (doseq [row rows]
    (add-row! sheet row)))


(defn add-sheet! 
  "Add a new worksheet to the workbook."
  [workbook name]
  (.createSheet workbook name))


(defn create-workbook
  "Create a new workbook with a single sheet and the data specified.
   The data is given a vector of vectors, representing the rows
   and the cells of the rows.

   For example, to create a workbook with a sheet with
   two rows of each three columns:

   (create-workbook \"Sheet 1\" 
                    [[\"Name\" \"Quantity\" \"Price\"]
                     [\"Foo Widget\" 2 42]])             
   "
  [sheet-name data]
  (let [workbook (XSSFWorkbook.)
	sheet    (add-sheet! workbook sheet-name)]
    (add-rows! sheet data)
    workbook))


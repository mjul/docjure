(ns dk.ative.docjure.spreadsheet
  (:import
    [java.io FileOutputStream FileInputStream InputStream OutputStream]
    [java.util Date Calendar]
    [org.apache.poi.xssf.usermodel XSSFWorkbook]
    [org.apache.poi.hssf.usermodel HSSFWorkbook]
    [org.apache.poi.ss.usermodel 
      Workbook Sheet Cell CellType Row
      Row$MissingCellPolicy
      HorizontalAlignment
      VerticalAlignment
      BorderStyle
      FillPatternType
      FormulaError
      WorkbookFactory DateUtil
      IndexedColors CellStyle Font
      CellValue Drawing CreationHelper]
    [org.apache.poi.ss.util CellReference AreaReference]))
;=

(defmacro assert-type [value expected-type]
  `(when-not (isa? (class ~value) ~expected-type)
     (throw (IllegalArgumentException.
             (format "%s is invalid. Expected %s. Actual type %s, value: %s"
                     (str '~value) ~expected-type (class ~value) ~value)))))

;; not used
(defn cell-reference [^Cell cell]
  (.formatAsString (CellReference. (.getRowIndex cell) (.getColumnIndex cell))))

(defmulti read-cell-value (fn [^CellValue cv _date-format?] (.getCellType cv)))
(defmethod read-cell-value CellType/BOOLEAN  [^CellValue cv _]  (.getBooleanValue cv))
(defmethod read-cell-value CellType/STRING   [^CellValue cv _]  (.getStringValue cv))
(defmethod read-cell-value CellType/NUMERIC  [^CellValue cv date-format?]
  (if date-format?
    (DateUtil/getJavaDate (.getNumberValue cv))
    (.getNumberValue cv)))
(defmethod read-cell-value CellType/ERROR    [^CellValue cv _]
  (keyword (.name (FormulaError/forInt (.getErrorValue cv)))))

(defmulti read-cell #(when % (.getCellType ^Cell %)))
(defmethod read-cell CellType/BLANK     [_]     nil)
(defmethod read-cell nil [_] nil)
(defmethod read-cell CellType/STRING    [^Cell cell]  (.getStringCellValue cell))
(defmethod read-cell CellType/FORMULA   [^Cell cell]
  (let [evaluator (.. cell getSheet getWorkbook
                      getCreationHelper createFormulaEvaluator)
        cv (.evaluate evaluator cell)]
    (if (and (= CellType/NUMERIC (.getCellType cv))
             (DateUtil/isCellDateFormatted cell))
      (.getDateCellValue cell)
      (read-cell-value cv false))))
(defmethod read-cell CellType/BOOLEAN   [^Cell cell]  (.getBooleanCellValue cell))
(defmethod read-cell CellType/NUMERIC   [^Cell cell]
  (if (DateUtil/isCellDateFormatted cell)
    (.getDateCellValue cell)
    (.getNumericCellValue cell)))
(defmethod read-cell CellType/ERROR     [^Cell cell]
  (keyword (.name (FormulaError/forInt (.getErrorCellValue cell)))))


(defn load-workbook-from-stream
  "Load an Excel workbook from a stream.
  The caller is required to close the stream after loading is completed."
  [^InputStream stream]
  (WorkbookFactory/create stream))

(defn load-workbook-from-file
  "Load an Excel .xls or .xlsx workbook from a file."
  [^String filename]
  (with-open [stream (FileInputStream. filename)]
    (load-workbook-from-stream stream)))

(defn load-workbook-from-resource
  "Load an Excel workbook from a named resource.
  Used when reading from a resource on a classpath
  as in the case of running on an application server."
  [^String resource]
  (let [url (clojure.java.io/resource resource)]
    (with-open [stream (.openStream url)]
      (load-workbook-from-stream stream))))

(defmulti load-workbook "Load an Excel .xls or .xlsx workbook from an InputStream." class)

(defmethod load-workbook String
  [filename]
  (load-workbook-from-file filename))

(defmethod load-workbook InputStream
  [stream]
  (load-workbook-from-stream stream))

(defn save-workbook-into-stream!
  "Save the workbook into a stream.
  The caller is required to close the stream after saving is completed."
  [^OutputStream stream ^Workbook workbook]
  (assert-type workbook Workbook)
  (let [rc (.write workbook stream)]
    (.flush stream)
    rc))

(defn save-workbook-into-file!
  "Save the workbook into a file."
  [^String filename ^Workbook workbook]
  (assert-type workbook Workbook)
  (with-open [file-out (FileOutputStream. filename)]
    (.write workbook file-out)))

(defmulti save-workbook!
          "Save the workbook into a stream or a file.
          In the case of saving into a stream, the caller is required
          to close the stream after saving is completed."
          (fn [x _] (class x)))

(defmethod save-workbook! OutputStream
  [stream workbook]
  (save-workbook-into-stream! stream workbook))

(defmethod save-workbook! String
  [filename workbook]
  (save-workbook-into-file! filename workbook))

(defn sheet-seq
  "Return a lazy seq of the sheets in a workbook."
  [^Workbook workbook]
  (assert-type workbook Workbook)
  (for [idx (range (.getNumberOfSheets workbook))]
    (.getSheetAt workbook idx)))

(defn sheet-name
  "Return the name of a sheet."
  [^Sheet sheet]
  (assert-type sheet Sheet)
  (.getSheetName sheet))

(defn- find-sheet
  [matching-fn ^Workbook workbook]
  (assert-type workbook Workbook)
  (->> (sheet-seq workbook)
       (filter matching-fn)
       first))

(defmulti select-sheet
  "Select a sheet from the workbook by name, regex or arbitrary predicate"
  (fn [predicate ^Workbook _workbook]
    (class predicate)))

(defmethod select-sheet String
  [name ^Workbook workbook]
  (find-sheet #(= name (sheet-name %)) workbook))

(defmethod select-sheet java.util.regex.Pattern
  [regex-pattern ^Workbook workbook]
  (find-sheet #(re-find regex-pattern (sheet-name %)) workbook))

(defmethod select-sheet :default
  [matching-fn ^Workbook workbook]
  (find-sheet matching-fn workbook))

(defn row-seq
  "Return a lazy sequence of the rows in a sheet. Missing rows will be returned as nil
  filter with e.g. (remove nil? (row-seq ...)) if missing rows are not needed"
  [^Sheet sheet]
  (assert-type sheet Sheet)
  (map #(.getRow sheet %)
       (range 0 (inc (.getLastRowNum sheet)))))

(defn- cell-seq-dispatch [x]
  (cond
   (isa? (class x) Row) :row
   (isa? (class x) Sheet) :sheet
   (seq? x) :coll
   :else :default))

(defmulti cell-seq
  "Return a seq of the cells in the input which can be a sheet, a row, or a collection
   of one of these. The seq is ordered ordered by sheet, row and column.
   Missing cells will be returned as nil, note this is different from blank cells which have type (CellType/BLANK)"
  cell-seq-dispatch)
(defmethod cell-seq :row  [^Row row] (map
                                       #(.getCell row % Row$MissingCellPolicy/RETURN_NULL_AND_BLANK)
                                       (range 0 (.getLastCellNum row))))
(defmethod cell-seq :sheet [sheet] (for [row (remove nil? (row-seq sheet))
                                         cell (cell-seq row)]
                                     cell))
(defmethod cell-seq :coll [coll] (for [x (remove nil? coll),
                                       cell (cell-seq x)]
                                   cell))

(defn into-seq
  [^Iterable sheet-or-row]
  (vec (for [item (iterator-seq (.iterator sheet-or-row))] item)))

(defn- project-cell [column-map ^Cell cell]
  (let [colname (-> cell
                    .getColumnIndex
                    org.apache.poi.ss.util.CellReference/convertNumToColString
                    keyword)
        new-key (column-map colname)]
    (when new-key
      {new-key (read-cell cell)})))

(defn select-columns
  "Takes two arguments: column hashmap and a sheet. The column hashmap
   specifies the mapping from spreadsheet columns dictionary keys:
   its keys are the spreadsheet column names and the values represent
   the names they are mapped to in the result.

   For example, to select columns A and C as :first and :third from the sheet

   (select-columns {:A :first, :C :third} sheet)
   => [{:first \"Value in cell A1\", :third \"Value in cell C1\"} ...] "
  [column-map ^Sheet sheet]
  (assert-type sheet Sheet)
  (vec
   (for [row (into-seq sheet)]
     (->> (map #(project-cell column-map %) row)
          (apply merge)))))

(defn string-cell? [^Cell cell]
  (= CellType/STRING (.getCellType cell)))

; (defn- date-or-calendar? [value]
;   (let [cls (class value)]
;     (or (isa? cls Date) (isa? cls Calendar))))

(defn- ^:dynamic create-date-format [^Workbook workbook ^String format]
  (let [date-style (.createCellStyle workbook)
        format-helper (.getCreationHelper workbook)]
    (doto date-style
          (.setDataFormat (.. format-helper createDataFormat (getFormat format))))))

(defn apply-date-format! [^Cell cell ^String format]
  (let [workbook (.. cell getSheet getWorkbook)
        date-style (.createCellStyle workbook)
        format-helper (.getCreationHelper workbook)]
    (.setDataFormat date-style
                    (.. format-helper createDataFormat (getFormat format)))
    (.setCellStyle cell date-style)))

(defmulti set-cell! (fn [^Cell _cell val] (type val)))

(defmethod set-cell! String [^Cell cell val]
  (when (= (.getCellType cell) CellType/FORMULA) 
    (.setCellType cell CellType/STRING))
  (.setCellValue cell ^String val))

(defmethod set-cell! Number [^Cell cell val]
  (when (= (.getCellType cell) CellType/FORMULA) 
    (.setCellType cell CellType/NUMERIC))
  (.setCellValue cell (double val)))

(defmethod set-cell! Boolean [^Cell cell val]
  (when (= (.getCellType cell) CellType/FORMULA) 
    (.setCellType cell CellType/BOOLEAN))
  (.setCellValue cell ^Boolean val))

(defmethod set-cell! Date [^Cell cell val]
  (when (= (.getCellType cell) CellType/FORMULA) 
    (.setCellType cell CellType/NUMERIC))
  (.setCellValue cell ^Date val)
  (.setCellStyle cell (create-date-format (.. cell getSheet getWorkbook) "m/d/yy")))

(defmethod set-cell! nil [^Cell cell _val]
  (let [^String null nil]
    (when (= (.getCellType cell) CellType/FORMULA) 
      (.setCellType cell CellType/BLANK))
    (.setCellValue cell null)))

(defn add-row! [^Sheet sheet values]
  (assert-type sheet Sheet)
  (let [row-num (if (= 0 (.getPhysicalNumberOfRows sheet))
                  0
                  (inc (.getLastRowNum sheet)))
        row (.createRow sheet row-num)]
    (doseq [[column-index value] (map-indexed #(list %1 %2) values)]
      (set-cell! (.createCell row column-index) value))
    row))

(defn add-rows!
  "Add rows to the sheet. The rows is a sequence of row-data, where
   each row-data is a sequence of values for the columns in increasing
   order on that row."
  [^Sheet sheet rows]
  (assert-type sheet Sheet)
  (binding [create-date-format (memoize create-date-format)]
    (doseq [row rows]
      (add-row! sheet row))))

(defn add-sheet!
  "Add a new sheet to the workbook."
  [^Workbook workbook name]
  (assert-type workbook Workbook)
  (.createSheet workbook name))

(defn create-workbook
  "Create a new XLSX workbook.  Sheet-name is a string name for the sheet. Data
  is a vector of vectors, representing the rows and the cells of the rows.
  Alternate sheet names and data to create multiple sheets.

  (create-workbook \"SheetName1\" [[\"A1\" \"A2\"][\"B1\" \"B2\"]]
                   \"SheetName2\" [[\"A1\" \"A2\"][\"B1\" \"B2\"]] "
 ([sheet-name data]
  (let [workbook (XSSFWorkbook.)
        sheet    (add-sheet! workbook sheet-name)]
    (add-rows! sheet data)
    workbook))

 ([sheet-name data & name-data-pairs]
  ;; incomplete pairs should not be allowed
  {:pre [(even? (count name-data-pairs))]}
  ;; call single arity version to create workbook
  (let [workbook (create-workbook sheet-name data)]
     ;; iterate through pairs adding sheets and rows
   (doseq [[s-name data] (partition 2 name-data-pairs)]
     (-> workbook
         (add-sheet! s-name)
         (add-rows!  data)))
   workbook)))

(defn- add-row-indexed!
  "Add row to the sheet, at a specific row index"
  [^Sheet sheet index values]
  (assert-type sheet Sheet)
  (let [row (.createRow sheet index)]
    (doseq [[column-index value] (map-indexed #(list %1 %2) values)]
      (when value                                             ; nil values are skipped
        (set-cell! (.createCell row column-index) value)))
    row))

(defn- add-sparse-rows!
  "Add rows to the sheet. rows is a sequence of row-data, where
   each row-data is a sequence of values for the columns in increasing
   order on that row., or nil to skip a row"
  [^Sheet sheet rows]
  (assert-type sheet Sheet)
  (doall
    (map-indexed (fn [index row]
                   (when-not (nil? row)
                     (add-row-indexed! sheet index row)))
                 rows)))

(defn create-sparse-workbook
  "Create a new XLSX workbook.  Sheet-name is a string name for the sheet. Data
  is a vector of vectors, representing the rows and the cells of the rows.
  Alternate sheet names and data to create multiple sheets.

  Spreadsheet rows and cells can be nil, which will create a sparse spreadsheet with
  non-continuous rows and cells

  (This version exists mostly for generating test data, `create-workbook` will
  normally do just fine unless you have a specific need for sparseness)

  (create-sparse-workbook \"SheetName1\" [[\"A1\" \"A2\"] nil [\"C1\" nil \"C3\"]]
                          \"SheetName2\" [[\"A1\" \"A2\"] nil [\"C1\" nil \"C3\"]] "
  ([sheet-name data]
   (let [workbook (XSSFWorkbook.)
         sheet    (add-sheet! workbook sheet-name)]
     (add-sparse-rows! sheet data)
     workbook))

  ([sheet-name data & name-data-pairs]
    ;; incomplete pairs should not be allowed
   {:pre [(even? (count name-data-pairs))]}
    ;; call single arity version to create workbook
   (let [workbook (create-sparse-workbook sheet-name data)]
     ;; iterate through pairs adding sheets and rows
     (doseq [[s-name data] (partition 2 name-data-pairs)]
       (-> workbook
           (add-sheet! s-name)
           (add-sparse-rows!  data)))
     workbook)))

(defn create-xls-workbook
  "Create a new XLS workbook with a single sheet and the data specified."
  [sheet-name data]
  (let [workbook (HSSFWorkbook.)
        sheet    (add-sheet! workbook sheet-name)]
    (add-rows! sheet data)
    workbook))

;******************************************************
;       helpers for font and style creation


(defn color-index
  "Returns color index from org.apache.ss.usermodel.IndexedColors
   from lowercase keywords"
  [colorkw]
  (.getIndex (IndexedColors/valueOf (.toUpperCase (name colorkw)))))

(defn horiz-align
  "Returns horizontal alignment"
  [kw]
  (case kw
    :left HorizontalAlignment/LEFT
    :right HorizontalAlignment/RIGHT
    :center HorizontalAlignment/CENTER))

(defn vert-align
  "Returns vertical alignment"
  [kw]
  (case kw
    :top VerticalAlignment/TOP
    :bottom VerticalAlignment/BOTTOM
    :center VerticalAlignment/CENTER))

(defn border
  "Returns border style"
  [kw]
  (case kw
    :thin BorderStyle/THIN
    :medium BorderStyle/MEDIUM
    :thick BorderStyle/THICK))

(defmacro whens
  "Processes any and all expressions whose tests evaluate to true.
   Example:
   (let [m (java.util.HashMap.)]
    (whens
     false (.put m :z 0)
     true  (.put m :a 1)
     true  (.put m :b 2)
     nil   (.put m :w 3))
    m)
   => {:b=2, :a=1}
  "
  [& [test expr :as clauses]]
  (when clauses
    `(do (when ~test ~expr)
         (whens ~@(nnext clauses)))))

;****************************************************

(defn create-font!
  "Create a new font in the workbook with options:

       :name      font family (string)
       :size      font size  (integer)
       :color     font color (keyword)
       :bold      true | false
       :italic    true | false
       :underline true | false

   Example:

      (create-font! wb
       {:name \"Arial\", :size 12, :color :blue,
        :bold true, :underline true})
   "
  [^Workbook workbook options]
  (assert-type workbook Workbook)
  (let [f (.createFont workbook)
        {:keys [name size color bold italic underline]} options]
    (whens
     name      (.setFontName f name)
     size      (.setFontHeightInPoints f size)
     color     (.setColor f (color-index color))
     bold      (.setBold f true)
     italic    (.setItalic f true)
     underline (.setUnderline f Font/U_SINGLE))
    f))

(defprotocol IFontable
  "A protocol that allows:
   1. interchangeable use of fonts and maps of font options
   2. getting fonts from either XLS or XLSX cell styles, which
      normally requires distinct syntax."
  (set-font [this style workbook])
  (get-font [this workbook])
  (as-font [this workbook]))

(extend-protocol IFontable
  clojure.lang.PersistentArrayMap
  (set-font [this ^CellStyle style workbook]
    (.setFont style (create-font! workbook this)))
  (as-font [this workbook] (create-font! workbook this))
  org.apache.poi.ss.usermodel.Font
  (set-font [this ^CellStyle style _] (.setFont style this))
  (as-font [this _] this)
  org.apache.poi.xssf.usermodel.XSSFCellStyle
  (get-font [this _] (.getFont this))
  org.apache.poi.hssf.usermodel.HSSFCellStyle
  (get-font [this workbook] (.getFont this workbook)))

(defn create-cell-style!
  "Create a new cell-style in the workbook from options:

      :background          background colour (as keyword)
      :font                font | fontmap (of font options)
      :halign              :left | :right | :center
      :valign              :top | :bottom | :center
      :wrap                true | false - controls text wrapping
      :border-left         :thin | :medium | :thick
      :border-right        :thin | :medium | :thick
      :border-top          :thin | :medium | :thick
      :border-bottom       :thin | :medium | :thick
      :left-border-color   colour keyword
      :right-border-color  colour keyword
      :top-border-color    colour keyword
      :bottom-border-color colour keyword
      :indent              number from 0 to 15
      :data-format         string

   Valid colour keywords are the colour names defined in
   org.apache.ss.usermodel.IndexedColors as lowercase keywords, eg.

     :black, :white, :red, :blue, :light_green, :yellow, ...

   Examples:
   I.
   (def f (create-font! wb {:name \"Arial\", :bold true, :italic true})
   (create-cell-style! wb {:background :yellow, :font f, :halign :center,
                           :wrap true, :border-bottom :thin})
   II.
   (create-cell-style! wb {:background :yellow, :halign :center,
                           :font {:name \"Arial\" :bold true :italic true},
                           :wrap true, :border-bottom :thin})
  "
  ([^Workbook workbook] (create-cell-style! workbook {}))

  ([^Workbook workbook styles]
   (assert-type workbook Workbook)
   (let [cs (.createCellStyle workbook)
         {:keys [background font halign valign wrap
                 border-left border-right border-top border-bottom
                 left-border-color right-border-color
                 top-border-color bottom-border-color
                 _borders indent data-format]} styles]
     (whens
      font   (set-font font cs workbook)
      background (do (.setFillForegroundColor cs (color-index background))
                     (.setFillPattern cs FillPatternType/SOLID_FOREGROUND))
      halign (.setAlignment cs (horiz-align halign))
      valign (.setVerticalAlignment cs (vert-align valign))
      wrap   (.setWrapText cs true)
      border-left (.setBorderLeft cs (border border-left))
      border-right (.setBorderRight cs (border border-right))
      border-top (.setBorderTop cs (border border-top))
      border-bottom (.setBorderBottom cs (border border-bottom))
      left-border-color (.setLeftBorderColor
                          cs (color-index left-border-color))
      right-border-color (.setRightBorderColor
                           cs (color-index right-border-color))
      top-border-color (.setTopBorderColor
                         cs (color-index top-border-color))
      bottom-border-color (.setBottomBorderColor
                            cs (color-index bottom-border-color))
      indent (.setIndention cs (short indent))
      data-format (let [df (.createDataFormat workbook)]
                    (.setDataFormat cs (.getFormat df data-format))))
     cs)))

(defn set-cell-style!
  "Apply a style to a cell.
   See also: create-cell-style!.
  "
  [^Cell cell ^CellStyle style]
  (assert-type cell Cell)
  (assert-type style CellStyle)
  (.setCellStyle cell style)
  cell)

(defn set-cell-comment!
  "Creates a cell comment-box that displays a comment string
   when the cell is hovered over. Returns the cell.

   Options:

   :font   (font | fontmap - font applied to the comment string)
   :width  (int - width of comment-box in columns; default 1 cols)
   :height (int - height of comment-box in rows; default 2 rows)

   Example:

   (set-cell-comment! acell \"This comment should\nspan two lines.\"
                     :width 2 :font {:bold true :size 12 :color blue})
   "
  [^Cell cell comment-str & {:keys [font width height]
                             :or {width 1, height 2}}]
  (let [sheet (.getSheet cell)
        wb (.getWorkbook sheet)
        drawing (.createDrawingPatriarch sheet)
        helper (.getCreationHelper wb)
        anchor (.createClientAnchor helper)
        c1 (.getColumnIndex cell)
        c2 (+ c1 width)
        r1 (.getRowIndex cell)
        r2 (+ r1 height)]
    (doto anchor
      (.setCol1 c1) (.setCol2 c2) (.setRow1 r1) (.setRow2 r2))
    (let [comment (.createCellComment drawing anchor)
          rts (.createRichTextString helper comment-str)]
      (when font
        (let [^Font f (as-font font wb)] (.applyFont rts f)))
      (.setString comment rts)
      (.setCellComment cell comment))
    cell))

(defn set-row-style!
  "Apply a style to all the cells in a row.
   Returns the row."
  [^Row row ^CellStyle style]
  (assert-type row Row)
  (assert-type style CellStyle)
  (doseq [^Cell c (cell-seq row)
          :when c]
    (.setCellStyle c style))
  row)

(defn get-row-styles
  "Returns a seq of the row's CellStyles.
  Missing cells will return a nil style"
  [^Row row]
  (map #(when % (.getCellStyle ^Cell %)) (cell-seq row)))

(defn set-row-styles!
  "Apply a seq of styles to the cells in a row.
  Cells that are missing won't be assigned a style - if you want to style missing cells, create them first"
  [^Row row styles]
  (let [pairs (map list (cell-seq row) styles)]
    (doseq [[^Cell c s] pairs]
      (when c (.setCellStyle c s)))))

(defn row-vec
  "Transform the row struct (hash-map) to a row vector according to the column order.
   Example:

     (row-vec [:foo :bar] {:foo \"Foo text\", :bar \"Bar text\"})
     > [\"Foo text\" \"Bar text\"]
  "
  [column-order row]
  (vec (map row column-order)))

(defn remove-row!
  "Remove a row from the sheet. Rows are not shifted up - the removed row will display as blank"
  [^Sheet sheet ^Row row]
  (assert-type sheet Sheet)
  (assert-type row Row)
  (.removeRow sheet row)
  sheet)

(defn remove-all-rows!
  "Remove all the rows from the sheet."
  [sheet]
  (doall
   (for [row (doall (remove nil? (row-seq sheet)))]
     (remove-row! sheet row)))
  sheet)

(defn- named-area-ref [^Workbook workbook n]
  (let [index (.getNameIndex workbook (name n))]
    (if (>= index 0)
      (-> (.getNameAt workbook index)
          (.getRefersToFormula)
          (AreaReference. (.getSpreadsheetVersion workbook)))
      nil)))

(defn- cell-from-ref [^Workbook workbook ^CellReference cref]
  (let [row (.getRow cref)
        col (int (.getCol cref))
        sheet (->> cref (.getSheetName) (.getSheet workbook))]
    (-> sheet (.getRow row) (.getCell col))))

(defn select-name
  "Given a workbook and name (string or keyword) of a named range, select-name
   returns a seq of cells or nil if the name could not be found."
  [^Workbook workbook n]
  (when-let [^AreaReference aref (named-area-ref workbook n)]
    (map (partial cell-from-ref workbook) (.getAllReferencedCells aref))))

(defn select-cell
  "Given a Sheet and a cell reference (A1), select-cell returns the cell
  or nil if the cell could not be found"
  [n ^Sheet sheet]
  (let [cellref (CellReference. n)
        row (.getRow cellref)
        col (.getCol cellref)]
    (try 
      (.getCell (.getRow sheet row) col) 
      (catch Exception _e nil))))

(defn add-name! [^Workbook workbook n string-ref]
  (let [the-name (.createName workbook)]
    (.setNameName the-name (name n))
    (.setRefersToFormula the-name string-ref)))

(defn cell-fn
  "Turn a cell (ideally containing a formula) into a function. The returned function
  will take a variable number of parameters, updating each of the inputcells in the
  sheet with the supplied values and return the value of the cell outputcell.
  Cell names are specified using Excel syntax, i.e. A2 or B12."
  [outputcell ^Sheet sheet & inputcells]
  (fn [& input] 
    (doseq [pair (seq (apply hash-map (interleave inputcells input)))]
      (set-cell! (select-cell (first pair) sheet) (last pair)))
    (read-cell (select-cell outputcell sheet))))

;;.

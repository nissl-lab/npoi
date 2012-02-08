/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.xssf.usermodel;






using NPOI.ss.formula.ptg.Ptg;
using NPOI.ss.formula.SharedFormula;
using NPOI.ss.formula.Eval.ErrorEval;
using NPOI.ss.SpreadsheetVersion;
using NPOI.ss.formula.FormulaParser;
using NPOI.ss.formula.FormulaRenderer;
using NPOI.ss.formula.FormulaType;
using NPOI.ss.usermodel.Cell;
using NPOI.ss.usermodel.CellStyle;
using NPOI.ss.usermodel.Comment;
using NPOI.ss.usermodel.DataFormatter;
using NPOI.ss.usermodel.DateUtil;
using NPOI.ss.usermodel.FormulaError;
using NPOI.ss.usermodel.Hyperlink;
using NPOI.ss.usermodel.RichTextString;
using NPOI.ss.util.CellRangeAddress;
using NPOI.ss.util.CellReference;
using NPOI.xssf.Model.SharedStringsTable;
using NPOI.xssf.Model.StylesTable;
using NPOI.util.Internal;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;

/**
 * High level representation of a cell in a row of a spreadsheet.
 * <p>
 * Cells can be numeric, formula-based or string-based (text).  The cell type
 * specifies this.  String cells cannot conatin numbers and numeric cells cannot
 * contain strings (at least according to our model).  Client apps should do the
 * conversions themselves.  Formula cells have the formula string, as well as
 * the formula result, which can be numeric or string.
 * </p>
 * <p>
 * Cells should have their number (0 based) before being Added to a row.  Only
 * cells that have values should be Added.
 * </p>
 */
public class XSSFCell : Cell {

    private static String FALSE_AS_STRING = "0";
    private static String TRUE_AS_STRING  = "1";

    /**
     * the xml bean Containing information about the cell's location, value,
     * data type, formatting, and formula
     */
    private CTCell _cell;

    /**
     * the XSSFRow this cell belongs to
     */
    private XSSFRow _row;

    /**
     * 0-based column index
     */
    private int _cellNum;

    /**
     * Table of strings shared across this workbook.
     * If two cells contain the same string, then the cell value is the same index into SharedStringsTable
     */
    private SharedStringsTable _sharedStringSource;

    /**
     * Table of cell styles shared across all cells in a workbook.
     */
    private StylesTable _stylesSource;

    /**
     * Construct a XSSFCell.
     *
     * @param row the parent row.
     * @param cell the xml bean Containing information about the cell.
     */
    protected XSSFCell(XSSFRow row, CTCell cell) {
        _cell = cell;
        _row = row;
        if (cell.GetR() != null) {
            _cellNum = new CellReference(cell.GetR()).Col;
        }
        _sharedStringSource = row.Sheet.GetWorkbook().GetSharedStringSource();
        _stylesSource = row.Sheet.GetWorkbook().GetStylesSource();
    }

    /**
     * @return table of strings shared across this workbook
     */
    protected SharedStringsTable GetSharedStringSource() {
        return _sharedStringSource;
    }

    /**
     * @return table of cell styles shared across this workbook
     */
    protected StylesTable GetStylesSource() {
        return _stylesSource;
    }

    /**
     * Returns the sheet this cell belongs to
     *
     * @return the sheet this cell belongs to
     */
    public XSSFSheet Sheet {
        return Row.Sheet;
    }

    /**
     * Returns the row this cell belongs to
     *
     * @return the row this cell belongs to
     */
    public XSSFRow Row {
        return _row;
    }

    /**
     * Get the value of the cell as a bool.
     * <p>
     * For strings, numbers, and errors, we throw an exception. For blank cells we return a false.
     * </p>
     * @return the value of the cell as a bool
     * @throws InvalidOperationException if the cell type returned by {@link #getCellType()}
     *   is not CELL_TYPE_BOOLEAN, CELL_TYPE_BLANK or CELL_TYPE_FORMULA
     */
    public bool GetBooleanCellValue() {
        int cellType = GetCellType();
        switch(cellType) {
            case CELL_TYPE_BLANK:
                return false;
            case CELL_TYPE_BOOLEAN:
                return _cell.IsSetV() && TRUE_AS_STRING.Equals(_cell.GetV());
            case CELL_TYPE_FORMULA:
                //YK: should throw an exception if requesting bool value from a non-bool formula
                return _cell.IsSetV() && TRUE_AS_STRING.Equals(_cell.GetV());
            default:
                throw typeMismatch(CELL_TYPE_BOOLEAN, cellType, false);
        }
    }

    /**
     * Set a bool value for the cell
     *
     * @param value the bool value to Set this cell to.  For formulas we'll Set the
     *        precalculated value, for bools we'll Set its value. For other types we
     *        will change the cell to a bool cell and Set its value.
     */
    public void SetCellValue(bool value) {
        _cell.SetT(STCellType.B);
        _cell.SetV(value ? TRUE_AS_STRING : FALSE_AS_STRING);
    }

    /**
     * Get the value of the cell as a number.
     * <p>
     * For strings we throw an exception. For blank cells we return a 0.
     * For formulas or error cells we return the precalculated value;
     * </p>
     * @return the value of the cell as a number
     * @throws InvalidOperationException if the cell type returned by {@link #getCellType()} is CELL_TYPE_STRING
     * @exception NumberFormatException if the cell value isn't a parsable <code>double</code>.
     * @see DataFormatter for turning this number into a string similar to that which Excel would render this number as.
     */
    public double GetNumericCellValue() {
        int cellType = GetCellType();
        switch(cellType) {
            case CELL_TYPE_BLANK:
                return 0.0;
            case CELL_TYPE_FORMULA:
            case CELL_TYPE_NUMERIC:
                if(_cell.IsSetV()) {
                   try {
                      return Double.Parse(_cell.GetV());
                   } catch(NumberFormatException e) {
                      throw typeMismatch(CELL_TYPE_NUMERIC, CELL_TYPE_STRING, false);
                   }
                } else {
                   return 0.0;
                }
            default:
                throw typeMismatch(CELL_TYPE_NUMERIC, cellType, false);
        }
    }


    /**
     * Set a numeric value for the cell
     *
     * @param value  the numeric value to Set this cell to.  For formulas we'll Set the
     *        precalculated value, for numerics we'll Set its value. For other types we
     *        will change the cell to a numeric cell and Set its value.
     */
    public void SetCellValue(double value) {
        if(Double.IsInfInity(value)) {
            // Excel does not support positive/negative infInities,
            // rather, it gives a #DIV/0! error in these cases.
            _cell.SetT(STCellType.E);
            _cell.SetV(FormulaError.DIV0.GetString());
        } else if (Double.IsNaN(value)){
            // Excel does not support Not-a-Number (NaN),
            // instead it immediately generates an #NUM! error.
            _cell.SetT(STCellType.E);
            _cell.SetV(FormulaError.NUM.GetString());
        } else {
            _cell.SetT(STCellType.N);
            _cell.SetV(String.ValueOf(value));
        }
    }

    /**
     * Get the value of the cell as a string
     * <p>
     * For numeric cells we throw an exception. For blank cells we return an empty string.
     * For formulaCells that are not string Formulas, we throw an exception
     * </p>
     * @return the value of the cell as a string
     */
    public String GetStringCellValue() {
        XSSFRichTextString str = GetRichStringCellValue();
        return str == null ? null : str.GetString();
    }

    /**
     * Get the value of the cell as a XSSFRichTextString
     * <p>
     * For numeric cells we throw an exception. For blank cells we return an empty string.
     * For formula cells we return the pre-calculated value if a string, otherwise an exception
     * </p>
     * @return the value of the cell as a XSSFRichTextString
     */
    public XSSFRichTextString GetRichStringCellValue() {
        int cellType = GetCellType();
        XSSFRichTextString rt;
        switch (cellType) {
            case CELL_TYPE_BLANK:
                rt = new XSSFRichTextString("");
                break;
            case CELL_TYPE_STRING:
                if (_cell.GetT() == STCellType.INLINE_STR) {
                    if(_cell.IsSetIs()) {
                        //string is expressed directly in the cell defInition instead of implementing the shared string table.
                        rt = new XSSFRichTextString(_cell.GetIs());
                    } else if (_cell.IsSetV()) {
                        //cached result of a formula
                        rt = new XSSFRichTextString(_cell.GetV());
                    } else {
                        rt = new XSSFRichTextString("");
                    }
                } else if (_cell.GetT() == STCellType.STR) {
                    //cached formula value
                    rt = new XSSFRichTextString(_cell.IsSetV() ? _cell.GetV() : "");
                } else {
                    if (_cell.IsSetV()) {
                        int idx = Int32.ParseInt(_cell.GetV());
                        rt = new XSSFRichTextString(_sharedStringSource.GetEntryAt(idx));
                    }
                    else {
                        rt = new XSSFRichTextString("");
                    }
                }
                break;
            case CELL_TYPE_FORMULA:
                CheckFormulaCachedValueType(CELL_TYPE_STRING, GetBaseCellType(false));
                rt = new XSSFRichTextString(_cell.IsSetV() ? _cell.GetV() : "");
                break;
            default:
                throw typeMismatch(CELL_TYPE_STRING, cellType, false);
        }
        rt.SetStylesTableReference(_stylesSource);
        return rt;
    }

    private static void CheckFormulaCachedValueType(int expectedTypeCode, int cachedValueType) {
        if (cachedValueType != expectedTypeCode) {
            throw typeMismatch(expectedTypeCode, cachedValueType, true);
        }
	}

	/**
     * Set a string value for the cell.
     *
     * @param str value to Set the cell to.  For formulas we'll Set the formula
     * cached string result, for String cells we'll Set its value. For other types we will
     * change the cell to a string cell and Set its value.
     * If value is null then we will change the cell to a Blank cell.
     */
    public void SetCellValue(String str) {
        SetCellValue(str == null ? null : new XSSFRichTextString(str));
    }

    /**
     * Set a string value for the cell.
     *
     * @param str  value to Set the cell to.  For formulas we'll Set the 'pre-Evaluated result string,
     * for String cells we'll Set its value.  For other types we will
     * change the cell to a string cell and Set its value.
     * If value is null then we will change the cell to a Blank cell.
     */
    public void SetCellValue(RichTextString str) {
        if(str == null || str.GetString() == null){
            SetCellType(Cell.CELL_TYPE_BLANK);
            return;
        }
        int cellType = GetCellType();
        switch(cellType){
            case Cell.CELL_TYPE_FORMULA:
                _cell.SetV(str.GetString());
                _cell.SetT(STCellType.STR);
                break;
            default:
                if(_cell.GetT() == STCellType.INLINE_STR) {
                    //set the 'pre-Evaluated result
                    _cell.SetV(str.GetString());
                } else {
                    _cell.SetT(STCellType.S);
                    XSSFRichTextString rt = (XSSFRichTextString)str;
                    rt.SetStylesTableReference(_stylesSource);
                    int sRef = _sharedStringSource.AddEntry(rt.GetCTRst());
                    _cell.SetV(Int32.ToString(sRef));
                }
                break;
        }
    }

    /**
     * Return a formula for the cell, for example, <code>SUM(C4:E4)</code>
     *
     * @return a formula for the cell
     * @throws InvalidOperationException if the cell type returned by {@link #getCellType()} is not CELL_TYPE_FORMULA
     */
    public String GetCellFormula() {
        int cellType = GetCellType();
        if(cellType != CELL_TYPE_FORMULA) throw typeMismatch(CELL_TYPE_FORMULA, cellType, false);

        CTCellFormula f = _cell.GetF();
        if (isPartOfArrayFormulaGroup() && f == null) {
            XSSFCell cell = Sheet.GetFirstCellInArrayFormula(this);
            return cell.GetCellFormula();
        }
        if (f.GetT() == STCellFormulaType.SHARED) {
            return ConvertSharedFormula((int)f.GetSi());
        }
        return f.StringValue;
    }

    /**
     * Creates a non shared formula from the shared formula counterpart
     *
     * @param si Shared Group Index
     * @return non shared formula Created for the given shared formula and this cell
     */
    private String ConvertSharedFormula(int si){
        XSSFSheet sheet = Sheet;

        CTCellFormula f = sheet.GetSharedFormula(si);
        if(f == null) throw new InvalidOperationException(
                "Master cell of a shared formula with sid="+si+" was not found");

        String sharedFormula = f.StringValue;
        //Range of cells which the shared formula applies to
        String sharedFormulaRange = f.GetRef();

        CellRangeAddress ref = CellRangeAddress.ValueOf(sharedFormulaRange);

        int sheetIndex = sheet.GetWorkbook().GetSheetIndex(sheet);
        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(sheet.GetWorkbook());
        SharedFormula sf = new SharedFormula(SpreadsheetVersion.EXCEL2007);

        Ptg[] ptgs = FormulaParser.Parse(sharedFormula, fpb, FormulaType.CELL, sheetIndex);
        Ptg[] fmla = sf.ConvertSharedFormulas(ptgs,
                RowIndex - ref.FirstRow, ColumnIndex - ref.FirstColumn);
        return FormulaRenderer.ToFormulaString(fpb, fmla);
    }

    /**
     * Sets formula for this cell.
     * <p>
     * Note, this method only Sets the formula string and does not calculate the formula value.
     * To Set the precalculated value use {@link #setCellValue(double)} or {@link #setCellValue(String)}
     * </p>
     *
     * @param formula the formula to Set, e.g. <code>"SUM(C4:E4)"</code>.
     *  If the argument is <code>null</code> then the current formula is Removed.
     * @throws NPOI.ss.formula.FormulaParseException if the formula has incorrect syntax or is otherwise invalid
     * @throws InvalidOperationException if the operation is not allowed, for example,
     *  when the cell is a part of a multi-cell array formula
     */
    public void SetCellFormula(String formula) {
        if(isPartOfArrayFormulaGroup()){
            NotifyArrayFormulaChanging();
        }
        SetFormula(formula, FormulaType.CELL);
    }

    /* namespace */ void SetCellArrayFormula(String formula, CellRangeAddress range) {
        SetFormula(formula, FormulaType.ARRAY);
        CTCellFormula cellFormula = _cell.GetF();
        cellFormula.SetT(STCellFormulaType.ARRAY);
        cellFormula.SetRef(range.formatAsString());
    }

    private void SetFormula(String formula, int formulaType) {
        XSSFWorkbook wb = _row.Sheet.GetWorkbook();
        if (formula == null) {
            wb.onDeleteFormula(this);
            if(_cell.IsSetF()) _cell.unsetF();
            return;
        }

        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
        //validate through the FormulaParser
        FormulaParser.Parse(formula, fpb, formulaType, wb.GetSheetIndex(getSheet()));

        CTCellFormula f = CTCellFormula.Factory.newInstance();
        f.SetStringValue(formula);
        _cell.SetF(f);
        if(_cell.IsSetV()) _cell.unsetV();
    }

    /**
     * Returns column index of this cell
     *
     * @return zero-based column index of a column in a sheet.
     */
    public int ColumnIndex {
        return this._cellNum;
    }

    /**
     * Returns row index of a row in the sheet that Contains this cell
     *
     * @return zero-based row index of a row in the sheet that Contains this cell
     */
    public int RowIndex {
        return _row.GetRowNum();
    }

    /**
     * Returns an A1 style reference to the location of this cell
     *
     * @return A1 style reference to the location of this cell
     */
    public String GetReference() {
        return _cell.GetR();
    }

    /**
     * Return the cell's style.
     *
     * @return the cell's style.</code>
     */
    public XSSFCellStyle GetCellStyle() {
        XSSFCellStyle style = null;
        if(_stylesSource.GetNumCellStyles() > 0){
            long idx = _cell.IsSetS() ? _cell.GetS() : 0;
            style = _stylesSource.GetStyleAt((int)idx);
        }
        return style;
    }

    /**
     * Set the style for the cell.  The style should be an XSSFCellStyle Created/retreived from
     * the XSSFWorkbook.
     *
     * @param style  reference Contained in the workbook.
     * If the value is null then the style information is Removed causing the cell to used the default workbook style.
     */
    public void SetCellStyle(CellStyle style) {
        if(style == null) {
            if(_cell.IsSetS()) _cell.unsetS();
        } else {
            XSSFCellStyle xStyle = (XSSFCellStyle)style;
            xStyle.verifyBelongsToStylesSource(_stylesSource);

            long idx = _stylesSource.PutStyle(xStyle);
            _cell.SetS(idx);
        }
    }

    /**
     * Return the cell type.
     *
     * @return the cell type
     * @see Cell#CELL_TYPE_BLANK
     * @see Cell#CELL_TYPE_NUMERIC
     * @see Cell#CELL_TYPE_STRING
     * @see Cell#CELL_TYPE_FORMULA
     * @see Cell#CELL_TYPE_BOOLEAN
     * @see Cell#CELL_TYPE_ERROR
     */
    public int GetCellType() {

        if (_cell.GetF() != null || Sheet.IsCellInArrayFormulaContext(this)) {
            return CELL_TYPE_FORMULA;
        }

        return GetBaseCellType(true);
    }

    /**
     * Only valid for formula cells
     * @return one of ({@link #CELL_TYPE_NUMERIC}, {@link #CELL_TYPE_STRING},
     *     {@link #CELL_TYPE_BOOLEAN}, {@link #CELL_TYPE_ERROR}) depending
     * on the cached value of the formula
     */
    public int GetCachedFormulaResultType() {
        if (_cell.GetF() == null) {
            throw new InvalidOperationException("Only formula cells have cached results");
        }

        return GetBaseCellType(false);
    }

    /**
     * Detect cell type based on the "t" attribute of the CTCell bean
     */
    private int GetBaseCellType(bool blankCells) {
        switch (_cell.GetT().intValue()) {
            case STCellType.INT_B:
                return CELL_TYPE_BOOLEAN;
            case STCellType.INT_N:
                if (!_cell.IsSetV() && blankCells) {
                    // ooxml does have a separate cell type of 'blank'.  A blank cell Gets encoded as
                    // (either not present or) a numeric cell with no value Set.
                    // The formula Evaluator (and perhaps other clients of this interface) needs to
                    // distinguish blank values which sometimes Get translated into zero and sometimes
                    // empty string, depending on context
                    return CELL_TYPE_BLANK;
                }
                return CELL_TYPE_NUMERIC;
            case STCellType.INT_E:
                return CELL_TYPE_ERROR;
            case STCellType.INT_S: // String is in shared strings
            case STCellType.INT_INLINE_STR: // String is inline in cell
            case STCellType.INT_STR:
                 return CELL_TYPE_STRING;
            default:
                throw new InvalidOperationException("Illegal cell type: " + this._cell.GetT());
        }
    }

    /**
     * Get the value of the cell as a date.
     * <p>
     * For strings we throw an exception. For blank cells we return a null.
     * </p>
     * @return the value of the cell as a date
     * @throws InvalidOperationException if the cell type returned by {@link #getCellType()} is CELL_TYPE_STRING
     * @exception NumberFormatException if the cell value isn't a parsable <code>double</code>.
     * @see DataFormatter for formatting  this date into a string similar to how excel does.
     */
    public Date GetDateCellValue() {
        int cellType = GetCellType();
        if (cellType == CELL_TYPE_BLANK) {
            return null;
        }

        double value = GetNumericCellValue();
        bool date1904 = Sheet.GetWorkbook().IsDate1904();
        return DateUtil.GetJavaDate(value, date1904);
    }

    /**
     * Set a date value for the cell. Excel treats dates as numeric so you will need to format the cell as
     * a date.
     *
     * @param value  the date value to Set this cell to.  For formulas we'll Set the
     *        precalculated value, for numerics we'll Set its value. For other types we
     *        will change the cell to a numeric cell and Set its value.
     */
    public void SetCellValue(Date value) {
        bool date1904 = Sheet.GetWorkbook().IsDate1904();
        SetCellValue(DateUtil.GetExcelDate(value, date1904));
    }

    /**
     * Set a date value for the cell. Excel treats dates as numeric so you will need to format the cell as
     * a date.
     * <p>
     * This will Set the cell value based on the Calendar's timezone. As Excel
     * does not support timezones this means that both 20:00+03:00 and
     * 20:00-03:00 will be reported as the same value (20:00) even that there
     * are 6 hours difference between the two times. This difference can be
     * preserved by using <code>setCellValue(value.GetTime())</code> which will
     * automatically shift the times to the default timezone.
     * </p>
     *
     * @param value  the date value to Set this cell to.  For formulas we'll Set the
     *        precalculated value, for numerics we'll Set its value. For othertypes we
     *        will change the cell to a numeric cell and Set its value.
     */
    public void SetCellValue(Calendar value) {
        bool date1904 = Sheet.GetWorkbook().IsDate1904();
        SetCellValue( DateUtil.GetExcelDate(value, date1904 ));
    }

    /**
     * Returns the error message, such as #VALUE!
     *
     * @return the error message such as #VALUE!
     * @throws InvalidOperationException if the cell type returned by {@link #getCellType()} isn't CELL_TYPE_ERROR
     * @see FormulaError
     */
    public String GetErrorCellString() {
        int cellType = GetBaseCellType(true);
        if(cellType != CELL_TYPE_ERROR) throw typeMismatch(CELL_TYPE_ERROR, cellType, false);

        return _cell.GetV();
    }
    /**
     * Get the value of the cell as an error code.
     * <p>
     * For strings, numbers, and bools, we throw an exception.
     * For blank cells we return a 0.
     * </p>
     *
     * @return the value of the cell as an error code
     * @throws InvalidOperationException if the cell type returned by {@link #getCellType()} isn't CELL_TYPE_ERROR
     * @see FormulaError
     */
    public byte GetErrorCellValue() {
        String code = GetErrorCellString();
        if (code == null) {
            return 0;
        }

        return FormulaError.forString(code).GetCode();
    }

    /**
     * Set a error value for the cell
     *
     * @param errorCode the error value to Set this cell to.  For formulas we'll Set the
     *        precalculated value , for errors we'll Set
     *        its value. For other types we will change the cell to an error
     *        cell and Set its value.
     * @see FormulaError
     */
    public void SetCellErrorValue(byte errorCode) {
        FormulaError error = FormulaError.forInt(errorCode);
        SetCellErrorValue(error);
    }

    /**
     * Set a error value for the cell
     *
     * @param error the error value to Set this cell to.  For formulas we'll Set the
     *        precalculated value , for errors we'll Set
     *        its value. For other types we will change the cell to an error
     *        cell and Set its value.
     */
    public void SetCellErrorValue(FormulaError error) {
        _cell.SetT(STCellType.E);
        _cell.SetV(error.GetString());
    }

    /**
     * Sets this cell as the active cell for the worksheet.
     */
    public void SetAsActiveCell() {
        Sheet.SetActiveCell(_cell.GetR());
    }

    /**
     * Blanks this cell. Blank cells have no formula or value but may have styling.
     * This method erases all the data previously associated with this cell.
     */
    private void SetBlank(){
        CTCell blank = CTCell.Factory.newInstance();
        blank.SetR(_cell.GetR());
        if(_cell.IsSetS()) blank.SetS(_cell.GetS());
        _cell.Set(blank);
    }

    /**
     * Sets column index of this cell
     *
     * @param num column index of this cell
     */
    protected void SetCellNum(int num) {
        CheckBounds(num);
        _cellNum = num;
        String ref = new CellReference(getRowIndex(), ColumnIndex).formatAsString();
        _cell.SetR(ref);
    }

    /**
     * Set the cells type (numeric, formula or string)
     *
     * @throws ArgumentException if the specified cell type is invalid
     * @see #CELL_TYPE_NUMERIC
     * @see #CELL_TYPE_STRING
     * @see #CELL_TYPE_FORMULA
     * @see #CELL_TYPE_BLANK
     * @see #CELL_TYPE_BOOLEAN
     * @see #CELL_TYPE_ERROR
     */
    public void SetCellType(int cellType) {
        int prevType = GetCellType();
       
        if(isPartOfArrayFormulaGroup()){
            NotifyArrayFormulaChanging();
        }
        if(prevType == CELL_TYPE_FORMULA && cellType != CELL_TYPE_FORMULA) {
            Sheet.GetWorkbook().onDeleteFormula(this);
        }
        
        switch (cellType) {
            case CELL_TYPE_BLANK:
                SetBlank();
                break;
            case CELL_TYPE_BOOLEAN:
                String newVal = ConvertCellValueToBoolean() ? TRUE_AS_STRING : FALSE_AS_STRING;
                _cell.SetT(STCellType.B);
                _cell.SetV(newVal);
                break;
            case CELL_TYPE_NUMERIC:
                _cell.SetT(STCellType.N);
                break;
            case CELL_TYPE_ERROR:
                _cell.SetT(STCellType.E);
                break;
            case CELL_TYPE_STRING:
                if(prevType != CELL_TYPE_STRING){
                    String str = ConvertCellValueToString();
                    XSSFRichTextString rt = new XSSFRichTextString(str);
                    rt.SetStylesTableReference(_stylesSource);
                    int sRef = _sharedStringSource.AddEntry(rt.GetCTRst());
                    _cell.SetV(Int32.ToString(sRef));
                }
                _cell.SetT(STCellType.S);
                break;
            case CELL_TYPE_FORMULA:
                if(!_cell.IsSetF()){
                    CTCellFormula f =  CTCellFormula.Factory.newInstance();
                    f.SetStringValue("0");
                    _cell.SetF(f);
                    if(_cell.IsSetT()) _cell.unsetT();
                }
                break;
            default:
                throw new ArgumentException("Illegal cell type: " + cellType);
        }
        if (cellType != CELL_TYPE_FORMULA && _cell.IsSetF()) {
			_cell.unsetF();
		}
    }

    /**
     * Returns a string representation of the cell
     * <p>
     * Formula cells return the formula string, rather than the formula result.
     * Dates are displayed in dd-MMM-yyyy format
     * Errors are displayed as #ERR&lt;errIdx&gt;
     * </p>
     */
    public override String ToString() {
        switch (getCellType()) {
            case CELL_TYPE_BLANK:
                return "";
            case CELL_TYPE_BOOLEAN:
                return GetBooleanCellValue() ? "TRUE" : "FALSE";
            case CELL_TYPE_ERROR:
                return ErrorEval.GetText(getErrorCellValue());
            case CELL_TYPE_FORMULA:
                return GetCellFormula();
            case CELL_TYPE_NUMERIC:
                if (DateUtil.IsCellDateFormatted(this)) {
                    DateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy");
                    return sdf.format(getDateCellValue());
                }
                return GetNumericCellValue() + "";
            case CELL_TYPE_STRING:
                return GetRichStringCellValue().ToString();
            default:
                return "Unknown Cell Type: " + GetCellType();
        }
    }

    /**
     * Returns the raw, underlying ooxml value for the cell
     * <p>
     * If the cell Contains a string, then this value is an index into
     * the shared string table, pointing to the actual string value. Otherwise,
     * the value of the cell is expressed directly in this element. Cells Containing formulas express
     * the last calculated result of the formula in this element.
     * </p>
     *
     * @return the raw cell value as Contained in the underlying CTCell bean,
     *     <code>null</code> for blank cells.
     */
    public String GetRawValue() {
        return _cell.GetV();
    }

    /**
     * Used to help format error messages
     */
    private static String GetCellTypeName(int cellTypeCode) {
        switch (cellTypeCode) {
            case CELL_TYPE_BLANK:   return "blank";
            case CELL_TYPE_STRING:  return "text";
            case CELL_TYPE_BOOLEAN: return "bool";
            case CELL_TYPE_ERROR:   return "error";
            case CELL_TYPE_NUMERIC: return "numeric";
            case CELL_TYPE_FORMULA: return "formula";
        }
        return "#unknown cell type (" + cellTypeCode + ")#";
    }

    /**
     * Used to help format error messages
     */
    private static RuntimeException typeMismatch(int expectedTypeCode, int actualTypeCode, bool IsFormulaCell) {
        String msg = "Cannot Get a "
            + GetCellTypeName(expectedTypeCode) + " value from a "
            + GetCellTypeName(actualTypeCode) + " " + (isFormulaCell ? "formula " : "") + "cell";
        return new InvalidOperationException(msg);
    }

    /**
     * @throws RuntimeException if the bounds are exceeded.
     */
    private static void CheckBounds(int cellIndex) {
        SpreadsheetVersion v = SpreadsheetVersion.EXCEL2007;
        int maxcol = SpreadsheetVersion.EXCEL2007.GetLastColumnIndex();
        if (cellIndex < 0 || cellIndex > maxcol) {
            throw new ArgumentException("Invalid column index (" + cellIndex
                    + ").  Allowable column range for " + v.name() + " is (0.."
                    + maxcol + ") or ('A'..'" + v.GetLastColumnName() + "')");
        }
    }

    /**
     * Returns cell comment associated with this cell
     *
     * @return the cell comment associated with this cell or <code>null</code>
     */
    public XSSFComment GetCellComment() {
        return Sheet.GetCellComment(_row.GetRowNum(), ColumnIndex);
    }

    /**
     * Assign a comment to this cell. If the supplied comment is null,
     * the comment for this cell will be Removed.
     *
     * @param comment the XSSFComment associated with this cell
     */
    public void SetCellComment(Comment comment) {
        if(comment == null) {
            RemoveCellComment();
            return;
        }

        comment.SetRow(getRowIndex());
        comment.SetColumn(getColumnIndex());
    }

    /**
     * Removes the comment for this cell, if there is one.
    */
    public void RemoveCellComment() {
        XSSFComment comment = GetCellComment();
        if(comment != null){
            String ref = _cell.GetR();
            XSSFSheet sh = Sheet;
            sh.GetCommentsTable(false).RemoveComment(ref);
            sh.GetVMLDrawing(false).RemoveCommentShape(getRowIndex(), ColumnIndex);
        }
    }

    /**
     * Returns hyperlink associated with this cell
     *
     * @return hyperlink associated with this cell or <code>null</code> if not found
     */
    public XSSFHyperlink GetHyperlink() {
        return Sheet.GetHyperlink(_row.GetRowNum(), _cellNum);
    }

    /**
     * Assign a hypelrink to this cell
     *
     * @param hyperlink the hypelrink to associate with this cell
     */
    public void SetHyperlink(Hyperlink hyperlink) {
        XSSFHyperlink link = (XSSFHyperlink)hyperlink;

        // Assign to us
        link.SetCellReference( new CellReference(_row.GetRowNum(), _cellNum).formatAsString() );

        // Add to the lists
        Sheet.AddHyperlink(link);
    }

    /**
     * Returns the xml bean Containing information about the cell's location (reference), value,
     * data type, formatting, and formula
     *
     * @return the xml bean Containing information about this cell
     */
    
    public CTCell GetCTCell(){
        return _cell;
    }

    /**
     * Chooses a new bool value for the cell when its type is changing.<p/>
     *
     * Usually the caller is calling SetCellType() with the intention of calling
     * SetCellValue(bool) straight Afterwards.  This method only exists to give
     * the cell a somewhat reasonable value until the SetCellValue() call (if at all).
     * TODO - perhaps a method like SetCellTypeAndValue(int, Object) should be introduced to avoid this
     */
    private bool ConvertCellValueToBoolean() {
        int cellType = GetCellType();

        if (cellType == CELL_TYPE_FORMULA) {
            cellType = GetBaseCellType(false);
        }

        switch (cellType) {
            case CELL_TYPE_BOOLEAN:
                return TRUE_AS_STRING.Equals(_cell.GetV());
            case CELL_TYPE_STRING:
                int sstIndex = Int32.ParseInt(_cell.GetV());
                XSSFRichTextString rt = new XSSFRichTextString(_sharedStringSource.GetEntryAt(sstIndex));
                String text = rt.GetString();
                return Boolean.ParseBoolean(text);
            case CELL_TYPE_NUMERIC:
                return Double.Parse(_cell.GetV()) != 0;

            case CELL_TYPE_ERROR:
            case CELL_TYPE_BLANK:
                return false;
        }
        throw new RuntimeException("Unexpected cell type (" + cellType + ")");
    }

    private String ConvertCellValueToString() {
        int cellType = GetCellType();

        switch (cellType) {
            case CELL_TYPE_BLANK:
                return "";
            case CELL_TYPE_BOOLEAN:
                return TRUE_AS_STRING.Equals(_cell.GetV()) ? "TRUE" : "FALSE";
            case CELL_TYPE_STRING:
                int sstIndex = Int32.ParseInt(_cell.GetV());
                XSSFRichTextString rt = new XSSFRichTextString(_sharedStringSource.GetEntryAt(sstIndex));
                return rt.GetString();
            case CELL_TYPE_NUMERIC:
            case CELL_TYPE_ERROR:
                return _cell.GetV();
            case CELL_TYPE_FORMULA:
                // should really Evaluate, but HSSFCell can't call HSSFFormulaEvaluator
                // just use cached formula result instead
                break;
            default:
                throw new InvalidOperationException("Unexpected cell type (" + cellType + ")");
        }
        cellType = GetBaseCellType(false);
        String textValue = _cell.GetV();
        switch (cellType) {
            case CELL_TYPE_BOOLEAN:
                if (TRUE_AS_STRING.Equals(textValue)) {
                    return "TRUE";
                }
                if (FALSE_AS_STRING.Equals(textValue)) {
                    return "FALSE";
                }
                throw new InvalidOperationException("Unexpected bool cached formula value '"
                    + textValue + "'.");
            case CELL_TYPE_STRING:
            case CELL_TYPE_NUMERIC:
            case CELL_TYPE_ERROR:
                return textValue;
        }
        throw new InvalidOperationException("Unexpected formula result type (" + cellType + ")");
    }

    public CellRangeAddress GetArrayFormulaRange() {
        XSSFCell cell = Sheet.GetFirstCellInArrayFormula(this);
        if (cell == null) {
            throw new InvalidOperationException("Cell " + _cell.GetR()
                    + " is not part of an array formula.");
        }
        String formulaRef = cell._cell.GetF().GetRef();
        return CellRangeAddress.ValueOf(formulaRef);
    }

    public bool IsPartOfArrayFormulaGroup() {
        return Sheet.IsCellInArrayFormulaContext(this);
    }

    /**
     * The purpose of this method is to validate the cell state prior to modification
     *
     * @see #NotifyArrayFormulaChanging()
     */
    void NotifyArrayFormulaChanging(String msg){
        if(isPartOfArrayFormulaGroup()){
            CellRangeAddress cra = GetArrayFormulaRange();
            if(cra.GetNumberOfCells() > 1) {
                throw new InvalidOperationException(msg);
            }
            //un-register the Single-cell array formula from the parent XSSFSheet
            Row.Sheet.RemoveArrayFormula(this);
        }
    }

    /**
     * Called when this cell is modified.
     * <p>
     * The purpose of this method is to validate the cell state prior to modification.
     * </p>
     *
     * @see #setCellType(int)
     * @see #setCellFormula(String)
     * @see XSSFRow#RemoveCell(NPOI.ss.usermodel.Cell)
     * @see NPOI.xssf.usermodel.XSSFSheet#RemoveRow(NPOI.ss.usermodel.Row)
     * @see NPOI.xssf.usermodel.XSSFSheet#ShiftRows(int, int, int)
     * @see NPOI.xssf.usermodel.XSSFSheet#AddMergedRegion(NPOI.ss.util.CellRangeAddress)
     * @throws InvalidOperationException if modification is not allowed
     */
    void NotifyArrayFormulaChanging(){
        CellReference ref = new CellReference(this);
        String msg = "Cell "+ref.formatAsString()+" is part of a multi-cell array formula. " +
                "You cannot change part of an array.";
        NotifyArrayFormulaChanging(msg);
    }
}



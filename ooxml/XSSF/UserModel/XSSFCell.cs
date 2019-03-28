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

using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Util;
using System;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using NPOI.SS;
using NPOI.Util;
using NPOI.SS.Formula.Eval;
using System.Globalization;
namespace NPOI.XSSF.UserModel
{

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
    public class XSSFCell : ICell
    {

        private static String FALSE_AS_STRING = "0";
        private static String TRUE_AS_STRING = "1";

        /**
         * the xml bean Containing information about the cell's location, value,
         * data type, formatting, and formula
         */
        private CT_Cell _cell;

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
        public XSSFCell(XSSFRow row, CT_Cell cell)
        {
            _cell = cell;
            _row = row;
            if (cell.r != null)
            {
                _cellNum = new CellReference(cell.r).Col;
            }
            else
            {
                int prevNum = row.LastCellNum;
                if (prevNum != -1)
                {
                    _cellNum = (row as XSSFRow).GetCell(prevNum - 1, MissingCellPolicy.RETURN_NULL_AND_BLANK).ColumnIndex + 1;
                }
            }
            _sharedStringSource = ((XSSFWorkbook)row.Sheet.Workbook).GetSharedStringSource();
            _stylesSource = ((XSSFWorkbook)row.Sheet.Workbook).GetStylesSource();
        }

        /**
         * @return table of strings shared across this workbook
         */
        protected SharedStringsTable GetSharedStringSource()
        {
            return _sharedStringSource;
        }

        /**
         * @return table of cell styles shared across this workbook
         */
        protected StylesTable GetStylesSource()
        {
            return _stylesSource;
        }

        /**
         * Returns the sheet this cell belongs to
         *
         * @return the sheet this cell belongs to
         */
        public ISheet Sheet
        {
            get
            {
                return _row.Sheet;
            }
        }

        /**
         * Returns the row this cell belongs to
         *
         * @return the row this cell belongs to
         */
        public IRow Row
        {
            get
            {
                return _row;
            }
        }

        /**
         * Get the value of the cell as a bool.
         * <p>
         * For strings, numbers, and errors, we throw an exception. For blank cells we return a false.
         * </p>
         * @return the value of the cell as a bool
         * @throws InvalidOperationException if the cell type returned by {@link #CellType}
         *   is not CellType.Boolean, CellType.Blank or CellType.Formula
         */
        public bool BooleanCellValue
        {
            get
            {
                CellType cellType = CellType;
                switch (cellType)
                {
                    case CellType.Blank:
                        return false;
                    case CellType.Boolean:
                        return _cell.IsSetV() && TRUE_AS_STRING.Equals(_cell.v);
                    case CellType.Formula:
                        //YK: should throw an exception if requesting bool value from a non-bool formula
                        return _cell.IsSetV() && TRUE_AS_STRING.Equals(_cell.v);
                    default:
                        throw TypeMismatch(CellType.Boolean, cellType, false);
                }
            }
        }

        /**
         * Set a bool value for the cell
         *
         * @param value the bool value to Set this cell to.  For formulas we'll Set the
         *        precalculated value, for bools we'll Set its value. For other types we
         *        will change the cell to a bool cell and Set its value.
         */
        public void SetCellValue(bool value)
        {
            _cell.t = (ST_CellType.b);
            _cell.v = (value ? TRUE_AS_STRING : FALSE_AS_STRING);
        }

        /**
         * Get the value of the cell as a number.
         * <p>
         * For strings we throw an exception. For blank cells we return a 0.
         * For formulas or error cells we return the precalculated value;
         * </p>
         * @return the value of the cell as a number
         * @throws InvalidOperationException if the cell type returned by {@link #CellType} is CellType.String
         * @exception NumberFormatException if the cell value isn't a parsable <code>double</code>.
         * @see DataFormatter for turning this number into a string similar to that which Excel would render this number as.
         */
        public double NumericCellValue
        {
            get
            {
                CellType cellType = CellType;
                switch (cellType)
                {
                    case CellType.Blank:
                        return 0.0;
                    case CellType.Formula:
                    case CellType.Numeric:
                        if (_cell.IsSetV())
                        {
                            if (string.IsNullOrEmpty(_cell.v))
                                return 0.0;
                            try
                            {
                                return Double.Parse(_cell.v, CultureInfo.InvariantCulture);
                            }
                            catch (FormatException)
                            {
                                throw TypeMismatch(CellType.Numeric, CellType.String, false);
                            }
                        }
                        else
                        {
                            return 0.0;
                        }
                    default:
                        throw TypeMismatch(CellType.Numeric, cellType, false);
                }
            }
        }


        /**
         * Set a numeric value for the cell
         *
         * @param value  the numeric value to Set this cell to.  For formulas we'll Set the
         *        precalculated value, for numerics we'll Set its value. For other types we
         *        will change the cell to a numeric cell and Set its value.
         */
        public void SetCellValue(double value)
        {
            if (Double.IsInfinity(value))
            {
                // Excel does not support positive/negative infInities,
                // rather, it gives a #DIV/0! error in these cases.
                _cell.t = (ST_CellType.e);
                _cell.v = (FormulaError.DIV0.String);
            }
            else if (Double.IsNaN(value))
            {
                // Excel does not support Not-a-Number (NaN),
                // instead it immediately generates an #NUM! error.
                _cell.t = (ST_CellType.e);
                _cell.v = (FormulaError.NUM.String);
            }
            else
            {
                _cell.t = (ST_CellType.n);
                _cell.v = (value.ToString(CultureInfo.InvariantCulture));
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
        public String StringCellValue
        {
            get
            {
                IRichTextString str = this.RichStringCellValue;
                return str == null ? null : str.String;
            }
        }

        /**
         * Get the value of the cell as a XSSFRichTextString
         * <p>
         * For numeric cells we throw an exception. For blank cells we return an empty string.
         * For formula cells we return the pre-calculated value if a string, otherwise an exception
         * </p>
         * @return the value of the cell as a XSSFRichTextString
         */
        public IRichTextString RichStringCellValue
        {
            get
            {
                CellType cellType = CellType;
                XSSFRichTextString rt;
                switch (cellType)
                {
                    case CellType.Blank:
                        rt = new XSSFRichTextString("");
                        break;
                    case CellType.String:
                        if (_cell.t == ST_CellType.inlineStr)
                        {
                            if (_cell.IsSetIs())
                            {
                                //string is expressed directly in the cell defInition instead of implementing the shared string table.
                                rt = new XSSFRichTextString(_cell.@is);
                            }
                            else if (_cell.IsSetV())
                            {
                                //cached result of a formula
                                rt = new XSSFRichTextString(_cell.v);
                            }
                            else
                            {
                                rt = new XSSFRichTextString("");
                            }
                        }
                        else if (_cell.t == ST_CellType.str)
                        {
                            //cached formula value
                            rt = new XSSFRichTextString(_cell.IsSetV() ? _cell.v : "");
                        }
                        else
                        {
                            if (_cell.IsSetV())
                            {
                                int idx = Int32.Parse(_cell.v);
                                rt = new XSSFRichTextString(_sharedStringSource.GetEntryAt(idx));
                            }
                            else
                            {
                                rt = new XSSFRichTextString("");
                            }
                        }
                        break;
                    case CellType.Formula:
                        CheckFormulaCachedValueType(CellType.String, GetBaseCellType(false));
                        rt = new XSSFRichTextString(_cell.IsSetV() ? _cell.v : "");
                        break;
                    default:
                        throw TypeMismatch(CellType.String, cellType, false);
                }
                rt.SetStylesTableReference(_stylesSource);
                return rt;
            }
        }

        private static void CheckFormulaCachedValueType(CellType expectedTypeCode, CellType cachedValueType)
        {
            if (cachedValueType != expectedTypeCode)
            {
                throw TypeMismatch(expectedTypeCode, cachedValueType, true);
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
        public void SetCellValue(String str)
        {
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
        public void SetCellValue(IRichTextString str)
        {
            if (str == null || string.IsNullOrEmpty(str.String))
            {
                SetCellType(CellType.Blank);
                return;
            }

            if (str.Length > SpreadsheetVersion.EXCEL2007.MaxTextLength)
            {
                throw new ArgumentException("The maximum length of cell contents (text) is 32,767 characters");
            }
            CellType cellType = CellType;
            switch (cellType)
            {
                case CellType.Formula:
                    _cell.v = (str.String);
                    _cell.t= (ST_CellType.str);
                    break;
                default:
                    if (_cell.t == ST_CellType.inlineStr)
                    {
                        //set the 'pre-Evaluated result
                        _cell.v = str.String;
                    }
                    else
                    {
                        _cell.t = ST_CellType.s;
                        XSSFRichTextString rt = (XSSFRichTextString)str;
                        rt.SetStylesTableReference(_stylesSource);
                        int sRef = _sharedStringSource.AddEntry(rt.GetCTRst());
                        _cell.v=sRef.ToString();
                    }
                    break;
            }
        }

        /// <summary>
        /// Return a formula for the cell,  for example, <code>SUM(C4:E4)</code>
        /// </summary>
        public String CellFormula
        {
            get
            {
                CellType cellType = CellType;
                if (cellType != CellType.Formula) 
                    throw TypeMismatch(CellType.Formula, cellType, false);

                CT_CellFormula f = _cell.f;
                if (IsPartOfArrayFormulaGroup && f == null)
                {
                    ICell cell = ((XSSFSheet)Sheet).GetFirstCellInArrayFormula(this);
                    return cell.CellFormula;
                }
                if (f.t == ST_CellFormulaType.shared)
                {
                    return ConvertSharedFormula((int)f.si);
                }
                return f.Value;
            }
            set 
            {
                SetCellFormula(value);
            }
        }

        /// <summary>
        /// Creates a non shared formula from the shared formula counterpart
        /// </summary>
        /// <param name="si">Shared Group Index</param>
        /// <returns>non shared formula created for the given shared formula and this cell</returns>
        private String ConvertSharedFormula(int si)
        {
            XSSFSheet sheet = (XSSFSheet)Sheet;

            CT_CellFormula f = sheet.GetSharedFormula(si);
            if (f == null) throw new InvalidOperationException(
                     "Master cell of a shared formula with sid=" + si + " was not found");

            String sharedFormula = f.Value;
            //Range of cells which the shared formula applies to
            String sharedFormulaRange = f.@ref;

            CellRangeAddress ref1 = CellRangeAddress.ValueOf(sharedFormulaRange);

            int sheetIndex = sheet.Workbook.GetSheetIndex(sheet);
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(sheet.Workbook);
            SharedFormula sf = new SharedFormula(SpreadsheetVersion.EXCEL2007);

            Ptg[] ptgs = FormulaParser.Parse(sharedFormula, fpb, FormulaType.Cell, sheetIndex);
            Ptg[] fmla = sf.ConvertSharedFormulas(ptgs,
                    RowIndex - ref1.FirstRow, ColumnIndex - ref1.FirstColumn);
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
        public void SetCellFormula(String formula)
        {
            if (IsPartOfArrayFormulaGroup)
            {
                NotifyArrayFormulaChanging();
            }
            SetFormula(formula, FormulaType.Cell);
        }

        internal void SetCellArrayFormula(String formula, CellRangeAddress range)
        {
            SetFormula(formula, FormulaType.Array);
            CT_CellFormula cellFormula = _cell.f;
            cellFormula.t = (ST_CellFormulaType.array);
            cellFormula.@ref = (range.FormatAsString());
        }

        private void SetFormula(String formula, FormulaType formulaType)
        {
            IWorkbook wb = _row.Sheet.Workbook;
            if (formula == null)
            {
                ((XSSFWorkbook)wb).OnDeleteFormula(this);
                if (_cell.IsSetF()) _cell.unsetF();
                return;
            }

            IFormulaParsingWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            //validate through the FormulaParser
            FormulaParser.Parse(formula, fpb, formulaType, wb.GetSheetIndex(this.Sheet));

            CT_CellFormula f = new CT_CellFormula();
            f.Value = formula;
            _cell.f= (f);
            if (_cell.IsSetV()) _cell.unsetV();
        }

        /// <summary>
        /// Returns zero-based column index of this cell
        /// </summary>
        public int ColumnIndex
        {
            get
            {
                return this._cellNum;
            }
        }

        /// <summary>
        /// Returns zero-based row index of a row in the sheet that contains this cell
        /// </summary>
        public int RowIndex
        {
            get
            {
                return _row.RowNum;
            }
        }
        /// <summary>
        /// Returns an A1 style reference to the location of this cell
        /// </summary>
        /// <returns>A1 style reference to the location of this cell</returns>
        public String GetReference()
        {
            String ref1 = _cell.r;
            if (ref1 == null)
            {
                return new CellReference(this).FormatAsString();
            }
            return ref1;
        }

        /// <summary>
        /// Return the cell's style.
        /// </summary>
        public ICellStyle CellStyle
        {
            get
            {
                XSSFCellStyle style = null;
                if ((null != _stylesSource) && (_stylesSource.NumCellStyles > 0))
                {
                    long idx = _cell.IsSetS() ? _cell.s : 0;
                    style = _stylesSource.GetStyleAt((int)idx);
                }
                return style;
            }
            set 
            {
                if (value == null)
                {
                    if (_cell.IsSetS()) _cell.unsetS();
                }
                else
                {
                    XSSFCellStyle xStyle = (XSSFCellStyle)value;
                    xStyle.VerifyBelongsToStylesSource(_stylesSource);

                    long idx = _stylesSource.PutStyle(xStyle);
                    _cell.s = (uint)idx;
                }
            }
        }
        /// <summary>
        /// Return the cell type.
        /// </summary>
        public CellType CellType
        {
            get
            {

                if (_cell.f != null || ((XSSFSheet)Sheet).IsCellInArrayFormulaContext(this))
                {
                    return CellType.Formula;
                }

                return GetBaseCellType(true);
            }
        }
        /// <summary>
        /// Only valid for formula cells
        /// </summary>
        public CellType CachedFormulaResultType
        {
            get
            {
                if (_cell.f == null)
                {
                    throw new InvalidOperationException("Only formula cells have cached results");
                }

                return GetBaseCellType(false);
            }
        }

        /// <summary>
        /// Detect cell type based on the "t" attribute of the CT_Cell bean
        /// </summary>
        /// <param name="blankCells"></param>
        /// <returns></returns>
        private CellType GetBaseCellType(bool blankCells)
        {
            switch (_cell.t)
            {
                case ST_CellType.b:
                    return CellType.Boolean;
                case ST_CellType.n:
                    if (!_cell.IsSetV() && blankCells)
                    {
                        // ooxml does have a separate cell type of 'blank'.  A blank cell Gets encoded as
                        // (either not present or) a numeric cell with no value Set.
                        // The formula Evaluator (and perhaps other clients of this interface) needs to
                        // distinguish blank values which sometimes Get translated into zero and sometimes
                        // empty string, depending on context
                        return CellType.Blank;
                    }
                    return CellType.Numeric;
                case ST_CellType.e:
                    return CellType.Error;
                case ST_CellType.s: // String is in shared strings
                case ST_CellType.inlineStr: // String is inline in cell
                case ST_CellType.str:
                    return CellType.String;
                default:
                    throw new InvalidOperationException("Illegal cell type: " + this._cell.t);
            }
        }

        /// <summary>
        /// Get the value of the cell as a date.
        /// </summary>
        public DateTime DateCellValue
        {
            get
            {
                CellType cellType = CellType;
                if (cellType == CellType.Blank)
                {
                    return DateTime.MinValue;
                }

                double value = NumericCellValue;
                bool date1904 = ((XSSFWorkbook)Sheet.Workbook).IsDate1904();
                return DateUtil.GetJavaDate(value, date1904);
            }
        }

        /// <summary>
        ///  Set a date value for the cell. Excel treats dates as numeric so you will need to format the cell as a date.
        /// </summary>
        /// <param name="value">the date value to Set this cell to.  For formulas we'll set the precalculated value, 
        /// for numerics we'll Set its value. For other types we will change the cell to a numeric cell and Set its value. </param>
        public void SetCellValue(DateTime value)
        {
            bool date1904 = ((XSSFWorkbook)Sheet.Workbook).IsDate1904();
            SetCellValue(DateUtil.GetExcelDate(value, date1904));
        }
        /// <summary>
        /// Returns the error message, such as #VALUE!
        /// </summary>
        public String ErrorCellString
        {
            get
            {
                CellType cellType = GetBaseCellType(true);
                if (cellType != CellType.Error) throw TypeMismatch(CellType.Error, cellType, false);

                return _cell.v;
            }
        }

        /// <summary>
        /// Get the value of the cell as an error code.
        /// For strings, numbers, and bools, we throw an exception.
        /// For blank cells we return a 0.
        /// </summary>
        public byte ErrorCellValue
        {
            get
            {
                String code = this.ErrorCellString;
                if (code == null)
                {
                    return 0;
                }

                return FormulaError.ForString(code).Code;
            }
        }
        public void SetCellErrorValue(byte errorCode)
        {
            FormulaError error = FormulaError.ForInt(errorCode);
            SetCellErrorValue(error);
        }
        /// <summary>
        /// Set a error value for the cell
        /// </summary>
        /// <param name="error">the error value to Set this cell to. 
        /// For formulas we'll Set the precalculated value , for errors we'll set
        /// its value. For other types we will change the cell to an error cell and Set its value.
        /// </param>
        public void SetCellErrorValue(FormulaError error)
        {
            _cell.t = (ST_CellType.e);
            _cell.v = (error.String);
        }

        /// <summary>
        /// Sets this cell as the active cell for the worksheet.
        /// </summary>
        public void SetAsActiveCell()
        {
            ((XSSFSheet)Sheet).SetActiveCell(GetReference());
        }

        /// <summary>
        /// Blanks this cell. Blank cells have no formula or value but may have styling.
        /// This method erases all the data previously associated with this cell.
        /// </summary>
        private void SetBlank()
        {
            CT_Cell blank = new CT_Cell();
            blank.r = (_cell.r);
            if (_cell.IsSetS()) blank.s=(_cell.s);
            _cell.Set(blank);
        }

        /// <summary>
        /// Sets column index of this cell
        /// </summary>
        /// <param name="num"></param>
        internal void SetCellNum(int num)
        {
            CheckBounds(num);
            _cellNum = num;
            String ref1 = new CellReference(RowIndex, ColumnIndex).FormatAsString();
            _cell.r = (ref1);
        }
        /// <summary>
        /// Set the cells type (numeric, formula or string)
        /// </summary>
        /// <param name="cellType"></param>
        public void SetCellType(CellType cellType)
        {
            CellType prevType = CellType;

            if (IsPartOfArrayFormulaGroup)
            {
                NotifyArrayFormulaChanging();
            }
            if (prevType == CellType.Formula && cellType != CellType.Formula)
            {
                ((XSSFWorkbook)Sheet.Workbook).OnDeleteFormula(this);
            }

            switch (cellType)
            {
                case CellType.Blank:
                    SetBlank();
                    break;
                case CellType.Boolean:
                    String newVal = ConvertCellValueToBoolean() ? TRUE_AS_STRING : FALSE_AS_STRING;
                    _cell.t= (ST_CellType.b);
                    _cell.v= (newVal);
                    break;
                case CellType.Numeric:
                    _cell.t = (ST_CellType.n);
                    break;
                case CellType.Error:
                    _cell.t = (ST_CellType.e);
                    break;
                case CellType.String:
                    if (prevType != CellType.String)
                    {
                        String str = ConvertCellValueToString();
                        XSSFRichTextString rt = new XSSFRichTextString(str);
                        rt.SetStylesTableReference(_stylesSource);
                        int sRef = _sharedStringSource.AddEntry(rt.GetCTRst());
                        _cell.v= sRef.ToString();
                    }
                    _cell.t= (ST_CellType.s);
                    break;
                case CellType.Formula:
                    if (!_cell.IsSetF())
                    {
                        CT_CellFormula f = new CT_CellFormula();
                        f.Value = "0";
                        _cell.f = (f);
                        if (_cell.IsSetT()) _cell.unsetT();
                    }
                    break;
                default:
                    throw new ArgumentException("Illegal cell type: " + cellType);
            }
            if (cellType != CellType.Formula && _cell.IsSetF())
            {
                _cell.unsetF();
            }
        }
        /// <summary>
        /// Returns a string representation of the cell
        /// </summary>
        /// <returns>Formula cells return the formula string, rather than the formula result.
        /// Dates are displayed in dd-MMM-yyyy format
        /// Errors are displayed as #ERR&lt;errIdx&gt;
        /// </returns>
        public override String ToString()
        {
            switch (CellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.Error:
                    return ErrorEval.GetText(ErrorCellValue);
                case CellType.Formula:
                    return CellFormula;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(this))
                    {
                        FormatBase sdf = new SimpleDateFormat("dd-MMM-yyyy");
                        return sdf.Format(DateCellValue, CultureInfo.CurrentCulture);
                    }
                    return NumericCellValue.ToString();
                case CellType.String:
                    return RichStringCellValue.ToString();
                default:
                    return "Unknown Cell Type: " + CellType;
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
         * @return the raw cell value as Contained in the underlying CT_Cell bean,
         *     <code>null</code> for blank cells.
         */
        public String GetRawValue()
        {
            return _cell.v;
        }

        /// <summary>
        /// Used to help format error messages
        /// </summary>
        /// <param name="cellTypeCode"></param>
        /// <returns></returns>
        private static String GetCellTypeName(CellType cellTypeCode)
        {
            switch (cellTypeCode)
            {
                case CellType.Blank: return "blank";
                case CellType.String: return "text";
                case CellType.Boolean: return "bool";
                case CellType.Error: return "error";
                case CellType.Numeric: return "numeric";
                case CellType.Formula: return "formula";
            }
            return "#unknown cell type (" + cellTypeCode + ")#";
        }

        /**
         * Used to help format error messages
         */
        private static Exception TypeMismatch(CellType expectedTypeCode, CellType actualTypeCode, bool IsFormulaCell)
        {
            String msg = "Cannot get a "
                + GetCellTypeName(expectedTypeCode) + " value from a "
                + GetCellTypeName(actualTypeCode) + " " + (IsFormulaCell ? "formula " : "") + "cell";
            return new InvalidOperationException(msg);
        }

        /**
         * @throws RuntimeException if the bounds are exceeded.
         */
        private static void CheckBounds(int cellIndex)
        {
            SpreadsheetVersion v = SpreadsheetVersion.EXCEL2007;
            int maxcol = SpreadsheetVersion.EXCEL2007.LastColumnIndex;
            if (cellIndex < 0 || cellIndex > maxcol)
            {
                throw new ArgumentException("Invalid column index (" + cellIndex
                        + ").  Allowable column range for " + v.ToString() + " is (0.."
                        + maxcol + ") or ('A'..'" + v.LastColumnName + "')");
            }
        }

        /// <summary>
        ///  Returns cell comment associated with this cell
        /// </summary>
        public IComment CellComment
        {
            get
            {
                return Sheet.GetCellComment(_row.RowNum, ColumnIndex);
            }
            set 
            {
                if (value == null)
                {
                    RemoveCellComment();
                    return;
                }

                value.Row = (RowIndex);
                value.Column = (ColumnIndex);
            }
        }

        /// <summary>
        /// Removes the comment for this cell, if there is one.
        /// </summary>
        public void RemoveCellComment() {
            IComment comment = this.CellComment;
            if (comment != null)
            {
                String ref1 = GetReference();
                XSSFSheet sh = (XSSFSheet)Sheet;
                sh.GetCommentsTable(false).RemoveComment(ref1);
                sh.GetVMLDrawing(false).RemoveCommentShape(RowIndex, ColumnIndex);
            }
        }

        /// <summary>
        /// Get or set hyperlink associated with this cell
        /// If the supplied hyperlink is null on setting, the hyperlink for this cell will be removed.
        /// </summary>
        public IHyperlink Hyperlink
        {
            get
            {
                return ((XSSFSheet)Sheet).GetHyperlink(_row.RowNum, _cellNum);
            }
            set 
            {
                if (value == null)
                {
                    RemoveHyperlink();
                    return;
                }
                XSSFHyperlink link = (XSSFHyperlink)value;

                // Assign to us
                link.SetCellReference(new CellReference(_row.RowNum, _cellNum).FormatAsString());

                // Add to the lists
                ((XSSFSheet)Sheet).AddHyperlink(link);
            }
        }

        /**
         * Removes the hyperlink for this cell, if there is one.
         */
        public void RemoveHyperlink()
        {
            ((XSSFSheet)Sheet).RemoveHyperlink(_row.RowNum, _cellNum);
        }
        /**
         * Returns the xml bean containing information about the cell's location (reference), value,
         * data type, formatting, and formula
         *
         * @return the xml bean containing information about this cell
         */
        internal CT_Cell GetCTCell()
        {
            return _cell;
        }

        /**
         * Chooses a new bool value for the cell when its type is changing.<p/>
         *
         * Usually the caller is calling SetCellType() with the intention of calling
         * SetCellValue(bool) straight afterwards.  This method only exists to give
         * the cell a somewhat reasonable value until the SetCellValue() call (if at all).
         * TODO - perhaps a method like SetCellTypeAndValue(int, Object) should be introduced to avoid this
         */
        private bool ConvertCellValueToBoolean()
        {
            CellType cellType = CellType;

            if (cellType == CellType.Formula)
            {
                cellType = GetBaseCellType(false);
            }

            switch (cellType)
            {
                case CellType.Boolean:
                    return TRUE_AS_STRING.Equals(_cell.v);
                case CellType.String:
                    int sstIndex = Int32.Parse(_cell.v);
                    XSSFRichTextString rt = new XSSFRichTextString(_sharedStringSource.GetEntryAt(sstIndex));
                    String text = rt.String;
                    return Boolean.Parse(text);
                case CellType.Numeric:
                    return Double.Parse(_cell.v, CultureInfo.InvariantCulture) != 0;

                case CellType.Error:
                case CellType.Blank:
                    return false;
            }
            throw new RuntimeException("Unexpected cell type (" + cellType + ")");
        }

        private String ConvertCellValueToString()
        {
            CellType cellType = CellType;

            switch (cellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return TRUE_AS_STRING.Equals(_cell.v) ? "TRUE" : "FALSE";
                case CellType.String:
                    int sstIndex = Int32.Parse(_cell.v);
                    XSSFRichTextString rt = new XSSFRichTextString(_sharedStringSource.GetEntryAt(sstIndex));
                    return rt.String;
                case CellType.Numeric:
                case CellType.Error:
                    return _cell.v;
                case CellType.Formula:
                    // should really Evaluate, but HSSFCell can't call HSSFFormulaEvaluator
                    // just use cached formula result instead
                    break;
                default:
                    throw new InvalidOperationException("Unexpected cell type (" + cellType + ")");
            }
            cellType = GetBaseCellType(false);
            String textValue = _cell.v;
            switch (cellType)
            {
                case CellType.Boolean:
                    if (TRUE_AS_STRING.Equals(textValue))
                    {
                        return "TRUE";
                    }
                    if (FALSE_AS_STRING.Equals(textValue))
                    {
                        return "FALSE";
                    }
                    throw new InvalidOperationException("Unexpected bool cached formula value '"
                        + textValue + "'.");
                case CellType.String:
                case CellType.Numeric:
                case CellType.Error:
                    return textValue;
            }
            throw new InvalidOperationException("Unexpected formula result type (" + cellType + ")");
        }

        public CellRangeAddress ArrayFormulaRange
        {
            get
            {
                XSSFCell cell = ((XSSFSheet)Sheet).GetFirstCellInArrayFormula(this);
                if (cell == null)
                {
                    throw new InvalidOperationException("Cell " + GetReference()
                            + " is not part of an array formula.");
                }
                String formulaRef = cell._cell.f.@ref;
                return CellRangeAddress.ValueOf(formulaRef);
            }
        }

        public bool IsPartOfArrayFormulaGroup
        {
            get
            {
                return ((XSSFSheet)Sheet).IsCellInArrayFormulaContext(this);
            }
        }

        /**
         * The purpose of this method is to validate the cell state prior to modification
         *
         * @see #NotifyArrayFormulaChanging()
         */
        internal void NotifyArrayFormulaChanging(String msg)
        {
            if (IsPartOfArrayFormulaGroup)
            {
                CellRangeAddress cra = this.ArrayFormulaRange;
                if (cra.NumberOfCells > 1)
                {
                    throw new InvalidOperationException(msg);
                }
                //un-register the Single-cell array formula from the parent XSSFSheet
                Row.Sheet.RemoveArrayFormula(this);
            }
        }

        /// <summary>
        /// Called when this cell is modified.The purpose of this method is to validate the cell state prior to modification.
        /// </summary>
        /// <exception cref="InvalidOperationException">if modification is not allowed</exception>
        internal void NotifyArrayFormulaChanging()
        {
            CellReference ref1 = new CellReference(this);
            String msg = "Cell " + ref1.FormatAsString() + " is part of a multi-cell array formula. " +
                    "You cannot change part of an array.";
            NotifyArrayFormulaChanging(msg);
        }

        #region ICell Members


        public bool IsMergedCell
        {
            get {
                return this.Sheet.IsMergedRegion(new CellRangeAddress(this.RowIndex, this.RowIndex, this.ColumnIndex, this.ColumnIndex));
            }
            
        }

        #endregion


        public ICell CopyCellTo(int targetIndex)
        {
            return CellUtil.CopyCell(this.Row, this.ColumnIndex, targetIndex);
        }

        public CellType GetCachedFormulaResultTypeEnum()
        {
            throw new NotImplementedException();
        }
    }


}

/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Collections;
    using System.IO;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.SS.Formula;
    using System.Globalization;
    using System.Collections.Generic;

    /// <summary>
    /// High level representation of a cell in a row of a spReadsheet.
    /// Cells can be numeric, formula-based or string-based (text).  The cell type
    /// specifies this.  String cells cannot conatin numbers and numeric cells cannot
    /// contain strings (at least according to our model).  Client apps should do the
    /// conversions themselves.  Formula cells have the formula string, as well as
    /// the formula result, which can be numeric or string.
    /// Cells should have their number (0 based) before being Added to a row.  Only
    /// cells that have values should be Added.
    /// </summary>
    /// <remarks>
    /// @author  Andrew C. Oliver (acoliver at apache dot org)
    /// @author  Dan Sherman (dsherman at Isisph.com)
    /// @author  Brian Sanders (kestrel at burdell dot org) Active Cell support
    /// @author  Yegor Kozlov cell comments support
    /// </remarks>
    [Serializable]
    public class HSSFCell : ICell
    {

        public const short ENCODING_UNCHANGED = -1;
        public const short ENCODING_COMPRESSED_UNICODE = 0;
        public const short ENCODING_UTF_16 = 1;

        private CellType cellType;
        private HSSFRichTextString stringValue;
        // fix warning CS0414 "never used": private short encoding = ENCODING_UNCHANGED;
        private HSSFWorkbook book;
        private HSSFSheet _sheet;
        private CellValueRecordInterface _record;
        private IComment comment;


        private const string FILE_FORMAT_NAME = "BIFF8";
        public static readonly int LAST_COLUMN_NUMBER = SpreadsheetVersion.EXCEL97.LastColumnIndex; // 2^8 - 1
        private static readonly string LAST_COLUMN_NAME = SpreadsheetVersion.EXCEL97.LastColumnName;

        /// <summary>
        /// Creates new Cell - Should only be called by HSSFRow.  This Creates a cell
        /// from scratch.
        /// When the cell is initially Created it is Set to CellType.Blank. Cell types
        /// can be Changed/overwritten by calling SetCellValue with the appropriate
        /// type as a parameter although conversions from one type to another may be
        /// prohibited.
        /// </summary>
        /// <param name="book">Workbook record of the workbook containing this cell</param>
        /// <param name="sheet">Sheet record of the sheet containing this cell</param>
        /// <param name="row">the row of this cell</param>
        /// <param name="col">the column for this cell</param>
        public HSSFCell(HSSFWorkbook book, HSSFSheet sheet, int row, short col)
            : this(book, sheet, row, col, CellType.Blank)
        {
        }

        /// <summary>
        /// Creates new Cell - Should only be called by HSSFRow.  This Creates a cell
        /// from scratch.
        /// </summary>
        /// <param name="book">Workbook record of the workbook containing this cell</param>
        /// <param name="sheet">Sheet record of the sheet containing this cell</param>
        /// <param name="row">the row of this cell</param>
        /// <param name="col">the column for this cell</param>
        /// <param name="type">CellType.Numeric, CellType.String, CellType.Formula, CellType.Blank,
        /// CellType.Boolean, CellType.Error</param>
        public HSSFCell(HSSFWorkbook book, HSSFSheet sheet, int row, short col,
                           CellType type)
        {
            CheckBounds(col);
            cellType = CellType.Unknown; // Force 'SetCellType' to Create a first Record
            stringValue = null;
            this.book = book;
            this._sheet = sheet;

            short xfindex = sheet.Sheet.GetXFIndexForColAt(col);
            SetCellType(type, false, row, col, xfindex);
        }

        /// <summary>
        /// Creates an Cell from a CellValueRecordInterface.  HSSFSheet uses this when
        /// reading in cells from an existing sheet.
        /// </summary>
        /// <param name="book">Workbook record of the workbook containing this cell</param>
        /// <param name="sheet">Sheet record of the sheet containing this cell</param>
        /// <param name="cval">the Cell Value Record we wish to represent</param>
        public HSSFCell(HSSFWorkbook book, HSSFSheet sheet, CellValueRecordInterface cval)
        {
            _record = cval;
            cellType = DetermineType(cval);
            stringValue = null;
            this.book = book;
            this._sheet = sheet;
            switch (cellType)
            {
                case CellType.String:
                    stringValue = new HSSFRichTextString(book.Workbook, (LabelSSTRecord)cval);
                    break;

                case CellType.Blank:
                    break;

                case CellType.Formula:
                    stringValue = new HSSFRichTextString(((FormulaRecordAggregate)cval).StringValue);
                    break;
            }
            //ExtendedFormatRecord xf = book.Workbook.GetExFormatAt(cval.XFIndex);

            //CellStyle = new HSSFCellStyle((short)cval.XFIndex, xf, book);
        }

        /**
         * private constructor to prevent blank construction
         */
        private HSSFCell()
        {
        }

        /**
         * used internally -- given a cell value record, figure out its type
         */
        private CellType DetermineType(CellValueRecordInterface cval)
        {
            if (cval is FormulaRecordAggregate)
            {
                return CellType.Formula;
            }

            Record record = (Record)cval;
            int sid = record.Sid;

            switch (sid)
            {

                case NumberRecord.sid:
                    return CellType.Numeric;


                case BlankRecord.sid:
                    return CellType.Blank;


                case LabelSSTRecord.sid:
                    return CellType.String;


                case FormulaRecordAggregate.sid:
                    return CellType.Formula;


                case BoolErrRecord.sid:
                    BoolErrRecord boolErrRecord = (BoolErrRecord)record;

                    return (boolErrRecord.IsBoolean)
                             ? CellType.Boolean
                             : CellType.Error;

            }
            throw new Exception("Bad cell value rec (" + cval.GetType().Name + ")");
        }
        /// <summary>
        /// the Workbook that this Cell is bound to
        /// </summary>
        public InternalWorkbook BoundWorkbook
        {
            get
            {
                return book.Workbook;
            }
        }

        public ISheet Sheet
        {
            get
            {
                return this._sheet;
            }
        }
        /// <summary>
        /// the HSSFRow this cell belongs to
        /// </summary>
        public IRow Row
        {
            get
            {
                int rowIndex = this.RowIndex;
                return _sheet.GetRow(rowIndex);
            }
        }

        /// <summary>
        /// Set the cells type (numeric, formula or string)
        /// </summary>
        /// <param name="cellType">Type of the cell.</param>
        public void SetCellType(CellType cellType)
        {
            NotifyFormulaChanging();
            if (IsPartOfArrayFormulaGroup)
            {
                NotifyArrayFormulaChanging();
            }
            int row = _record.Row;
            int col = _record.Column;
            short styleIndex = _record.XFIndex;
            SetCellType(cellType, true, row, col, styleIndex);
        }

        /// <summary>
        /// Sets the cell type. The SetValue flag indicates whether to bother about
        /// trying to preserve the current value in the new record if one is Created.
        /// The SetCellValue method will call this method with false in SetValue
        /// since it will overWrite the cell value later
        /// </summary>
        /// <param name="cellType">Type of the cell.</param>
        /// <param name="setValue">if set to <c>true</c> [set value].</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="styleIndex">Index of the style.</param>
        private void SetCellType(CellType cellType, bool setValue, int row, int col, short styleIndex)
        {
            if (cellType > CellType.Error)
            {
                throw new Exception("I have no idea what type that Is!");
            }
            switch (cellType)
            {

                case CellType.Formula:
                    FormulaRecordAggregate frec = null;

                    if (cellType != this.cellType)
                    {
                        frec = _sheet.Sheet.RowsAggregate.CreateFormula(row, col);
                    }
                    else
                    {
                        frec = (FormulaRecordAggregate)_record;
                    }
                    frec.Column = col;
                    if (setValue)
                    {
                        frec.FormulaRecord.Value = NumericCellValue;
                    }
                    frec.XFIndex = styleIndex;
                    frec.Row = row;
                    _record = frec;
                    break;

                case CellType.Numeric:
                    NumberRecord nrec = null;

                    if (cellType != this.cellType)
                    {
                        nrec = new NumberRecord();
                    }
                    else
                    {
                        nrec = (NumberRecord)_record;
                    }
                    nrec.Column = col;
                    if (setValue)
                    {
                        nrec.Value = NumericCellValue;
                    }
                    nrec.XFIndex = styleIndex;
                    nrec.Row = row;
                    _record = nrec;
                    break;

                case CellType.String:
                    LabelSSTRecord lrec = null;

                    if (cellType != this.cellType)
                    {
                        lrec = new LabelSSTRecord();
                    }
                    else
                    {
                        lrec = (LabelSSTRecord)_record;
                    }
                    lrec.Column = col;
                    lrec.Row = row;
                    lrec.XFIndex = styleIndex;
                    if (setValue)
                    {
                        String str = ConvertCellValueToString();
                        int sstIndex = book.Workbook.AddSSTString(new UnicodeString(str));
                        lrec.SSTIndex = (sstIndex);
                        UnicodeString us = book.Workbook.GetSSTString(sstIndex);
                        stringValue = new HSSFRichTextString();
                        stringValue.UnicodeString = us;
                    }
                    _record = lrec;
                    break;

                case CellType.Blank:
                    BlankRecord brec = null;

                    if (cellType != this.cellType)
                    {
                        brec = new BlankRecord();
                    }
                    else
                    {
                        brec = (BlankRecord)_record;
                    }
                    brec.Column = col;

                    // During construction the cellStyle may be null for a Blank cell.
                    brec.XFIndex = styleIndex;
                    brec.Row = row;
                    _record = brec;
                    break;

                case CellType.Boolean:
                    BoolErrRecord boolRec = null;

                    if (cellType != this.cellType)
                    {
                        boolRec = new BoolErrRecord();
                    }
                    else
                    {
                        boolRec = (BoolErrRecord)_record;
                    }
                    boolRec.Column = col;
                    if (setValue)
                    {
                        boolRec.SetValue(ConvertCellValueToBoolean());
                    }
                    boolRec.XFIndex = styleIndex;
                    boolRec.Row = row;
                    _record = boolRec;
                    break;

                case CellType.Error:
                    BoolErrRecord errRec = null;

                    if (cellType != this.cellType)
                    {
                        errRec = new BoolErrRecord();
                    }
                    else
                    {
                        errRec = (BoolErrRecord)_record;
                    }
                    errRec.Column = col;
                    if (setValue)
                    {
                        errRec.SetValue((byte)HSSFErrorConstants.ERROR_VALUE);
                    }
                    errRec.XFIndex = styleIndex;
                    errRec.Row = row;
                    _record = errRec;
                    break;
            }
            if (cellType != this.cellType &&
                this.cellType != CellType.Unknown)  // Special Value to indicate an Uninitialized Cell
            {
                _sheet.Sheet.ReplaceValueRecord(_record);
            }
            this.cellType = cellType;
        }

        /// <summary>
        /// Get the cells type (numeric, formula or string)
        /// </summary>
        /// <value>The type of the cell.</value>
        public CellType CellType
        {
            get
            {
                return cellType;
            }
        }
        private String ConvertCellValueToString()
        {

            switch (cellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return ((BoolErrRecord)_record).BooleanValue ? "TRUE" : "FALSE";
                case CellType.String:
                    int sstIndex = ((LabelSSTRecord)_record).SSTIndex;
                    return book.Workbook.GetSSTString(sstIndex).String;
                case CellType.Numeric:
                    return NumberToTextConverter.ToText(((NumberRecord)_record).Value);
                case CellType.Error:
                    return HSSFErrorConstants.GetText(((BoolErrRecord)_record).ErrorValue);
                case CellType.Formula:
                    // should really evaluate, but Cell can't call HSSFFormulaEvaluator
                    // just use cached formula result instead
                    break;
                default:
                    throw new InvalidDataException("Unexpected cell type (" + cellType + ")");
            }
            FormulaRecordAggregate fra = ((FormulaRecordAggregate)_record);
            FormulaRecord fr = fra.FormulaRecord;
            switch (fr.CachedResultType)
            {
                case CellType.Boolean:
                    return fr.CachedBooleanValue ? "TRUE" : "FALSE";
                case CellType.String:
                    return fra.StringValue;
                case CellType.Numeric:
                    return NumberToTextConverter.ToText(fr.Value); 
                case CellType.Error:
                    return HSSFErrorConstants.GetText(fr.CachedErrorValue);
            }
            throw new InvalidDataException("Unexpected formula result type (" + cellType + ")");

        }
        /// <summary>
        /// Set a numeric value for the cell
        /// </summary>
        /// <param name="value">the numeric value to Set this cell to.  For formulas we'll Set the
        /// precalculated value, for numerics we'll Set its value. For other types we
        /// will Change the cell to a numeric cell and Set its value.</param>
        public void SetCellValue(double value)
        {
            if(double.IsInfinity(value))
            {
                // Excel does not support positive/negative infinities,
                // rather, it gives a #DIV/0! error in these cases.
                SetCellErrorValue(FormulaError.DIV0.Code);
            }
            else if (double.IsNaN(value))
            {
                // Excel does not support Not-a-Number (NaN),
                // instead it immediately generates a #NUM! error.
                SetCellErrorValue(FormulaError.NUM.Code);
            }
            else
            {
                int row = _record.Row;
                int col = _record.Column;
                short styleIndex = _record.XFIndex;

                switch (cellType)
                {
                    case CellType.Numeric:
                        ((NumberRecord)_record).Value = value;
                        break;
                    case CellType.Formula:
                        ((FormulaRecordAggregate)_record).SetCachedDoubleResult(value);
                        break;
                    default:
                        SetCellType(CellType.Numeric, false, row, col, styleIndex);
                        ((NumberRecord)_record).Value = value;
                        break;
                }
            }
        
        }

        /// <summary>
        /// Set a date value for the cell. Excel treats dates as numeric so you will need to format the cell as
        /// a date.
        /// </summary>
        /// <param name="value">the date value to Set this cell to.  For formulas we'll Set the
        /// precalculated value, for numerics we'll Set its value. For other types we
        /// will Change the cell to a numeric cell and Set its value.</param>
        public void SetCellValue(DateTime value)
        {
            SetCellValue(DateUtil.GetExcelDate(value, this.book.Workbook.IsUsing1904DateWindowing));
        }


        /// <summary>
        /// Set a string value for the cell. Please note that if you are using
        /// full 16 bit Unicode you should call SetEncoding() first.
        /// </summary>
        /// <param name="value">value to Set the cell to.  For formulas we'll Set the formula
        /// string, for String cells we'll Set its value.  For other types we will
        /// Change the cell to a string cell and Set its value.
        /// If value is null then we will Change the cell to a Blank cell.</param>
        public void SetCellValue(String value)
        {
            HSSFRichTextString str = new HSSFRichTextString(value);
            SetCellValue(str);
        }
        /**
 * set a error value for the cell
 *
 * @param errorCode the error value to set this cell to.  For formulas we'll set the
 *        precalculated value , for errors we'll set
 *        its value. For other types we will change the cell to an error
 *        cell and set its value.
 */
        public void SetCellErrorValue(byte errorCode)
        {
            int row = _record.Row;
            int col = _record.Column;
            short styleIndex = _record.XFIndex;
            switch (cellType)
            {

                case CellType.Error:
                    ((BoolErrRecord)_record).SetValue(errorCode);
                    break;
                case CellType.Formula:
                    ((FormulaRecordAggregate)_record).SetCachedErrorResult(errorCode);
                    break;
                default:
                    SetCellType(CellType.Error, false, row, col, styleIndex);
                    ((BoolErrRecord)_record).SetValue(errorCode);
                    break;
            }
        }
        /// <summary>
        /// Set a string value for the cell. Please note that if you are using
        /// full 16 bit Unicode you should call SetEncoding() first.
        /// </summary>
        /// <param name="value">value to Set the cell to.  For formulas we'll Set the formula
        /// string, for String cells we'll Set its value.  For other types we will
        /// Change the cell to a string cell and Set its value.
        /// If value is null then we will Change the cell to a Blank cell.</param>
        public void SetCellValue(IRichTextString value)
        {
            int row = _record.Row;
            int col = _record.Column;
            short styleIndex = _record.XFIndex;
            if (value == null)
            {
                NotifyFormulaChanging();
                SetCellType(CellType.Blank, false, row, col, styleIndex);
                return;
            }

            if (value.Length > NPOI.SS.SpreadsheetVersion.EXCEL97.MaxTextLength)
            {
                throw new ArgumentException("The maximum length of cell contents (text) is 32,767 characters");
            }
            if (cellType == CellType.Formula)
            {
                // Set the 'pre-Evaluated result' for the formula 
                // note - formulas do not preserve text formatting.
                FormulaRecordAggregate fr = (FormulaRecordAggregate)_record;
                fr.SetCachedStringResult(value.String);
                // Update our local cache to the un-formatted version
                stringValue = new HSSFRichTextString(value.String);
                return;
            }

            if (cellType != CellType.String)
            {
                SetCellType(CellType.String, false, row, col, styleIndex);
            }
            int index = 0;

            HSSFRichTextString hvalue = (HSSFRichTextString)value;
            UnicodeString str = hvalue.UnicodeString;
            index = book.Workbook.AddSSTString(str);
            ((LabelSSTRecord)_record).SSTIndex = index;
            stringValue = hvalue;
            stringValue.SetWorkbookReferences(book.Workbook, ((LabelSSTRecord)_record));
            stringValue.UnicodeString = book.Workbook.GetSSTString(index);
        }

        /**
 * Should be called any time that a formula could potentially be deleted.
 * Does nothing if this cell currently does not hold a formula
 */
        private void NotifyFormulaChanging()
        {
            if (_record is FormulaRecordAggregate)
            {
                ((FormulaRecordAggregate)_record).NotifyFormulaChanging();
            }
        }

        /// <summary>
        /// Gets or sets the cell formula.
        /// </summary>
        /// <value>The cell formula.</value>
        public String CellFormula
        {
            get
            {
                if (!(_record is FormulaRecordAggregate))
                    throw TypeMismatch(CellType.Formula, cellType, true);

                return HSSFFormulaParser.ToFormulaString(book, ((FormulaRecordAggregate)_record).FormulaTokens);
            }
            set
            {
                SetCellFormula(value);
            }
        }

        public void SetCellFormula(String formula)
        {
            if (IsPartOfArrayFormulaGroup)
            {
                NotifyArrayFormulaChanging();
            }
            int row = _record.Row;
            int col = _record.Column;
            short styleIndex = _record.XFIndex;

            if (string.IsNullOrEmpty(formula))
            {
                NotifyFormulaChanging();
                SetCellType(CellType.Blank, false, row, col, styleIndex);
                return;
            }
            int sheetIndex = book.GetSheetIndex(_sheet);
            Ptg[] ptgs = HSSFFormulaParser.Parse(formula, book, FormulaType.Cell, sheetIndex);

            SetCellType(CellType.Formula, false, row, col, styleIndex);
            FormulaRecordAggregate agg = (FormulaRecordAggregate)_record;
            FormulaRecord frec = agg.FormulaRecord;
            frec.Options = ((short)2);
            frec.Value = (0);

            //only set to default if there is no extended format index already set
            if (agg.XFIndex == (short)0)
            {
                agg.XFIndex = ((short)0x0f);
            }
            agg.SetParsedExpression(ptgs);
        }

        /// <summary>
        /// Get the value of the cell as a number.  For strings we throw an exception.
        /// For blank cells we return a 0.
        /// </summary>
        /// <value>The numeric cell value.</value>
        public double NumericCellValue
        {
            get
            {
                switch (cellType)
                {
                    case CellType.Blank:
                        return 0.0;

                    case CellType.Numeric:
                        return ((NumberRecord)_record).Value;
                    case CellType.Formula:
                        break;
                    default:
                        throw TypeMismatch(CellType.Numeric, cellType, false);
                }
                FormulaRecord fr = ((FormulaRecordAggregate)_record).FormulaRecord;
                CheckFormulaCachedValueType(CellType.Numeric, fr);
                return fr.Value;
            }
        }

        /// <summary>
        /// Used to help format error messages
        /// </summary>
        /// <param name="cellTypeCode">The cell type code.</param>
        /// <returns></returns>
        private String GetCellTypeName(CellType cellTypeCode)
        {
            switch (cellTypeCode)
            {
                case CellType.Blank: return "blank";
                case CellType.String: return "text";
                case CellType.Boolean: return "boolean";
                case CellType.Error: return "error";
                case CellType.Numeric: return "numeric";
                case CellType.Formula: return "formula";
            }
            return "#unknown cell type (" + cellTypeCode + ")#";
        }


        /// <summary>
        /// Types the mismatch.
        /// </summary>
        /// <param name="expectedTypeCode">The expected type code.</param>
        /// <param name="actualTypeCode">The actual type code.</param>
        /// <param name="isFormulaCell">if set to <c>true</c> [is formula cell].</param>
        /// <returns></returns>
        private Exception TypeMismatch(CellType expectedTypeCode, CellType actualTypeCode, bool isFormulaCell)
        {
            String msg = "Cannot get a "
                + GetCellTypeName(expectedTypeCode) + " value from a "
                + GetCellTypeName(actualTypeCode) + " " + (isFormulaCell ? "formula " : "") + "cell";
            return new InvalidOperationException(msg);
        }

        /// <summary>
        /// Checks the type of the formula cached value.
        /// </summary>
        /// <param name="expectedTypeCode">The expected type code.</param>
        /// <param name="fr">The fr.</param>
        private void CheckFormulaCachedValueType(CellType expectedTypeCode, FormulaRecord fr)
        {
            CellType cachedValueType = fr.CachedResultType;
            if (cachedValueType != expectedTypeCode)
            {
                throw TypeMismatch(expectedTypeCode, cachedValueType, true);
            }
        }


        /// <summary>
        /// Get the value of the cell as a date.  For strings we throw an exception.
        /// For blank cells we return a null.
        /// </summary>
        /// <value>The date cell value.</value>
        public DateTime DateCellValue
        {
            get
            {
                if (cellType == CellType.Blank)
                {
                    return DateTime.MaxValue;
                }
                if (cellType == CellType.String)
                {
                    throw new InvalidDataException(
                        "You cannot get a date value from a String based cell");
                }
                if (cellType == CellType.Boolean)
                {
                    throw new InvalidDataException(
                        "You cannot get a date value from a bool cell");
                }
                if (cellType == CellType.Error)
                {
                    throw new InvalidDataException(
                        "You cannot get a date value from an error cell");
                }
                double value = this.NumericCellValue;
                if (book.Workbook.IsUsing1904DateWindowing)
                {
                    return DateUtil.GetJavaDate(value, true);
                }
                else
                {
                    return DateUtil.GetJavaDate(value, false);
                }
            }
        }

        /// <summary>
        /// Get the value of the cell as a string - for numeric cells we throw an exception.
        /// For blank cells we return an empty string.
        /// For formulaCells that are not string Formulas, we return empty String
        /// </summary>
        /// <value>The string cell value.</value>
        public String StringCellValue
        {
            get
            {
                IRichTextString str = RichStringCellValue;
                return str.String;
            }
        }

        /// <summary>
        /// Get the value of the cell as a string - for numeric cells we throw an exception.
        /// For blank cells we return an empty string.
        /// For formulaCells that are not string Formulas, we return empty String
        /// </summary>
        /// <value>The rich string cell value.</value>
        public IRichTextString RichStringCellValue
        {
            get
            {
                switch (cellType)
                {
                    case CellType.Blank:
                        return new HSSFRichTextString("");
                    case CellType.String:
                        return stringValue;
                    case CellType.Formula:
                        break;
                    default:
                        throw TypeMismatch(CellType.String, cellType, false);
                }
                FormulaRecordAggregate fra = ((FormulaRecordAggregate)_record);
                CheckFormulaCachedValueType(CellType.String, fra.FormulaRecord);
                String strVal = fra.StringValue;
                return new HSSFRichTextString(strVal == null ? "" : strVal);
            }
        }

        /// <summary>
        /// Set a bool value for the cell
        /// </summary>
        /// <param name="value">the bool value to Set this cell to.  For formulas we'll Set the
        /// precalculated value, for bools we'll Set its value. For other types we
        /// will Change the cell to a bool cell and Set its value.</param>
        public void SetCellValue(bool value)
        {
            int row = _record.Row;
            int col = _record.Column;
            short styleIndex = _record.XFIndex;
            switch (cellType)
            {
                case CellType.Boolean:
                    ((BoolErrRecord)_record).SetValue(value);
                    break;
                case CellType.Formula:
                    ((FormulaRecordAggregate)_record).SetCachedBooleanResult(value);
                    break;
                default:
                    SetCellType(CellType.Boolean, false, row, col, styleIndex);
                    ((BoolErrRecord)_record).SetValue(value);
                    break;
            }
        }
        /// <summary>
        /// Chooses a new bool value for the cell when its type is changing.
        /// Usually the caller is calling SetCellType() with the intention of calling
        /// SetCellValue(bool) straight afterwards.  This method only exists to give
        /// the cell a somewhat reasonable value until the SetCellValue() call (if at all).
        /// TODO - perhaps a method like SetCellTypeAndValue(int, Object) should be introduced to avoid this
        /// </summary>
        /// <returns></returns>
        private bool ConvertCellValueToBoolean()
        {

            switch (cellType)
            {
                case CellType.Boolean:
                    return ((BoolErrRecord)_record).BooleanValue;
                case CellType.String:
                    int sstIndex = ((LabelSSTRecord)_record).SSTIndex;
                    String text = book.Workbook.GetSSTString(sstIndex).String;
                    return Convert.ToBoolean(text, CultureInfo.CurrentCulture);

                case CellType.Numeric:
                    return ((NumberRecord)_record).Value != 0;

                // All other cases Convert to false
                // These choices are not well justified.
                case CellType.Formula:
                    // use cached formula result if it's the right type: 
                    FormulaRecord fr = ((FormulaRecordAggregate)_record).FormulaRecord;
                    CheckFormulaCachedValueType(CellType.Boolean, fr);
                    return fr.CachedBooleanValue;
                // Other cases convert to false 
                // These choices are not well justified. 
                case CellType.Error:
                case CellType.Blank:
                    return false;
            }
            throw new Exception("Unexpected cell type (" + cellType + ")");
        }

        /// <summary>
        /// Get the value of the cell as a bool.  For strings, numbers, and errors, we throw an exception.
        /// For blank cells we return a false.
        /// </summary>
        /// <value><c>true</c> if [boolean cell value]; otherwise, <c>false</c>.</value>
        public bool BooleanCellValue
        {
            get
            {
                switch (cellType)
                {
                    case CellType.Blank:
                        return false;
                    case CellType.Boolean:
                        return ((BoolErrRecord)_record).BooleanValue;
                    case CellType.Formula:
                        break;
                    default:
                        throw TypeMismatch(CellType.Boolean, cellType, false);

                }
                FormulaRecord fr = ((FormulaRecordAggregate)_record).FormulaRecord;
                CheckFormulaCachedValueType(CellType.Boolean, fr);
                return fr.CachedBooleanValue;
            }
        }

        /// <summary>
        /// Get the value of the cell as an error code.  For strings, numbers, and bools, we throw an exception.
        /// For blank cells we return a 0.
        /// </summary>
        /// <value>The error cell value.</value>
        public byte ErrorCellValue
        {
            get
            {
                switch (cellType)
                {
                    case CellType.Error:
                        return ((BoolErrRecord)_record).ErrorValue;
                    case CellType.Formula:
                        break;
                    default:
                        throw TypeMismatch(CellType.Error, cellType, false);

                }
                FormulaRecord fr = ((FormulaRecordAggregate)_record).FormulaRecord;
                CheckFormulaCachedValueType(CellType.Error, fr);
                return (byte)fr.CachedErrorValue;
            }
        }

        /// <summary>
        /// Get the style for the cell.  This is a reference to a cell style contained in the workbook
        /// object.
        /// </summary>
        /// <value>The cell style.</value>
        public ICellStyle CellStyle
        {
            get
            {
                short styleIndex = _record.XFIndex;
                ExtendedFormatRecord xf = book.Workbook.GetExFormatAt(styleIndex);
                return new HSSFCellStyle(styleIndex, xf, book);
            }
            set
            {
                // A style of null means resetting back to the default style
                if (value == null)
                {
                    _record.XFIndex = ((short)0xf);
                    return;
                }
                // Verify it really does belong to our workbook
                ((HSSFCellStyle)value).VerifyBelongsToWorkbook(book);

                short styleIndex;
                if (((HSSFCellStyle)value).UserStyleName != null)
                {
                    styleIndex = ApplyUserCellStyle((HSSFCellStyle)value);
                }
                else
                {
                    styleIndex = value.Index;
                }

                // Change our cell record to use this style
                _record.XFIndex = styleIndex;
            }
        }
        /**
 * Applying a user-defined style (UDS) is special. Excel does not directly reference user-defined styles, but
 * instead create a 'proxy' ExtendedFormatRecord referencing the UDS as parent.
 *
 * The proceudre to apply a UDS is as follows:
 *
 * 1. search for a ExtendedFormatRecord with parentIndex == style.getIndex()
 *    and xfType ==  ExtendedFormatRecord.XF_CELL.
 * 2. if not found then create a new ExtendedFormatRecord and copy all attributes from the user-defined style
 *    and set the parentIndex to be style.getIndex()
 * 3. return the index of the ExtendedFormatRecord, this will be assigned to the parent cell record
 *
 * @param style  the user style to apply
 *
 * @return  the index of a ExtendedFormatRecord record that will be referenced by the cell
 */
        private short ApplyUserCellStyle(HSSFCellStyle style)
        {
            if (style.UserStyleName == null)
            {
                throw new ArgumentException("Expected user-defined style");
            }

            InternalWorkbook iwb = book.Workbook;
            short userXf = -1;
            int numfmt = iwb.NumExFormats;
            for (short i = 0; i < numfmt; i++)
            {
                ExtendedFormatRecord xf = iwb.GetExFormatAt(i);
                if (xf.XFType == ExtendedFormatRecord.XF_CELL && xf.ParentIndex == style.Index)
                {
                    userXf = i;
                    break;
                }
            }
            short styleIndex;
            if (userXf == -1)
            {
                ExtendedFormatRecord xfr = iwb.CreateCellXF();
                xfr.CloneStyleFrom(iwb.GetExFormatAt(style.Index));
                xfr.IndentionOptions = (short)0;
                xfr.XFType = (ExtendedFormatRecord.XF_CELL);
                xfr.ParentIndex = (style.Index);
                styleIndex = (short)numfmt;
            }
            else
            {
                styleIndex = userXf;
            }

            return styleIndex;
        }

        /// <summary>
        /// Should only be used by HSSFSheet and friends.  Returns the low level CellValueRecordInterface record
        /// </summary>
        /// <value>the cell via the low level api.</value>
        public CellValueRecordInterface CellValueRecord
        {
            get { return _record; }
        }

        /// <summary>
        /// Checks the bounds.
        /// </summary>
        /// <param name="cellIndex">The cell num.</param>
        /// <exception cref="Exception">if the bounds are exceeded.</exception>
        private void CheckBounds(int cellIndex)
        {
            if (cellIndex < 0 || cellIndex > LAST_COLUMN_NUMBER)
            {
                throw new ArgumentException("Invalid column index (" + cellIndex
                        + ").  Allowable column range for " + FILE_FORMAT_NAME + " is (0.."
                        + LAST_COLUMN_NUMBER + ") or ('A'..'" + LAST_COLUMN_NAME + "')");
            }
        }

        /// <summary>
        /// Sets this cell as the active cell for the worksheet
        /// </summary>
        public void SetAsActiveCell()
        {
            int row = _record.Row;
            int col = _record.Column;

            this._sheet.Sheet.SetActiveCell(row, col);
        }

        /// <summary>
        /// Returns a string representation of the cell
        /// This method returns a simple representation,
        /// anthing more complex should be in user code, with
        /// knowledge of the semantics of the sheet being Processed.
        /// Formula cells return the formula string,
        /// rather than the formula result.
        /// Dates are Displayed in dd-MMM-yyyy format
        /// Errors are Displayed as #ERR&lt;errIdx&gt;
        /// </summary>
        public override String ToString()
        {
            switch (CellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.Error:
                    return NPOI.SS.Formula.Eval.ErrorEval.GetText(((BoolErrRecord)_record).ErrorValue);
                case CellType.Formula:
                    return CellFormula;
                case CellType.Numeric:
                    string format = this.CellStyle.GetDataFormatString();
                    DataFormatter formatter = new DataFormatter();
                    return formatter.FormatCellValue(this);
                case CellType.String:
                    return StringCellValue;
                default:
                    return "Unknown Cell Type: " + CellType;
            }

        }


        /// <summary>
        /// Returns comment associated with this cell
        /// </summary>
        /// <value>The cell comment associated with this cell.</value>
        public IComment CellComment
        {
            get
            {
                if (comment == null)
                {
                    comment = _sheet.FindCellComment(_record.Row, _record.Column);
                }
                return comment;
            }
            set
            {
                if (value == null)
                {
                    RemoveCellComment();
                    return;
                }

                value.Row = _record.Row;
                value.Column = _record.Column;
                this.comment = value;
            }
        }

        /// <summary>
        /// Removes the comment for this cell, if
        /// there is one.
        /// </summary>
        /// <remarks>WARNING - some versions of excel will loose
        /// all comments after performing this action!</remarks>
        public void RemoveCellComment()
        {
            HSSFComment comment2 = _sheet.FindCellComment(_record.Row, _record.Column);
            comment = null;
            if (null == comment2)
            {
                return;
            }
            (_sheet.DrawingPatriarch as HSSFPatriarch).RemoveShape(comment2);
        }

        /// <summary>
        /// Gets the index of the column.
        /// </summary>
        /// <value>The index of the column.</value>
        public int ColumnIndex
        {
            get
            {
                return _record.Column & 0xFFFF;
            }
        }
        /**
         * Updates the cell record's idea of what
         *  column it belongs in (0 based)
         * @param num the new cell number
         */
        internal void UpdateCellNum(int num)
        {
            _record.Column = num;
        }
        /// <summary>
        /// Gets the (zero based) index of the row containing this cell
        /// </summary>
        /// <value>The index of the row.</value>
        public int RowIndex
        {
            get
            {
                return _record.Row;
            }
        }
        /// <summary>
        /// Get or set hyperlink associated with this cell
        /// If the supplied hyperlink is null on setting, the hyperlink for this cell will be removed.
        /// </summary>
        /// <value>The hyperlink associated with this cell or null if not found</value>
        public IHyperlink Hyperlink
        {
            get
            {
                for (IEnumerator it = _sheet.Sheet.Records.GetEnumerator(); it.MoveNext(); )
                {
                    RecordBase rec = (RecordBase)it.Current;
                    if (rec is HyperlinkRecord)
                    {
                        HyperlinkRecord link = (HyperlinkRecord)rec;
                        if (link.FirstColumn == _record.Column && link.FirstRow == _record.Row)
                        {
                            return new HSSFHyperlink(link);
                        }
                    }
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    RemoveHyperlink();
                    return;
                }
                HSSFHyperlink link = (HSSFHyperlink)value;
                value.FirstRow = _record.Row;
                value.LastRow = _record.Row;
                value.FirstColumn = _record.Column;
                value.LastColumn = _record.Column;

                switch (link.Type)
                {
                    case HyperlinkType.Email:
                    case HyperlinkType.Url:
                        value.Label = ("url");
                        break;
                    case HyperlinkType.File:
                        value.Label = ("file");
                        break;
                    case HyperlinkType.Document:
                        value.Label = ("place");
                        break;
                }

                int eofLoc = _sheet.Sheet.FindFirstRecordLocBySid(EOFRecord.sid);
                _sheet.Sheet.Records.Insert(eofLoc, link.record);
            }
        }

        /// <summary>
        /// Removes the hyperlink for this cell, if there is one.
        /// </summary>
        public void RemoveHyperlink()
        {
            RecordBase toRemove = null;
            for (IEnumerator<RecordBase> it = _sheet.Sheet.Records.GetEnumerator(); it.MoveNext(); )
            {
                RecordBase rec = it.Current;
                if (rec is HyperlinkRecord)
                {
                    HyperlinkRecord link = (HyperlinkRecord)rec;
                    if (link.FirstColumn == _record.Column && link.FirstRow == _record.Row)
                    {
                        toRemove = rec;
                        break;
                        //it.Remove();
                        //return;
                    }
                }
            }
            if (toRemove != null)
                _sheet.Sheet.Records.Remove(toRemove);
        }

        /// <summary>
        /// Only valid for formula cells
        /// </summary>
        /// <value>one of (CellType.Numeric,CellType.String, CellType.Boolean, CellType.Error) depending
        /// on the cached value of the formula</value>
        public CellType CachedFormulaResultType
        {
            get
            {
                if (this.cellType != CellType.Formula)
                {
                    throw new InvalidOperationException("Only formula cells have cached results");
                }
                return ((FormulaRecordAggregate)_record).FormulaRecord.CachedResultType;
            }
        }
        public bool IsPartOfArrayFormulaGroup
        {
            get
            {
                if (cellType != CellType.Formula)
                {
                    return false;
                }
                return ((FormulaRecordAggregate)_record).IsPartOfArrayFormula;
            }
        }

        internal void SetCellArrayFormula(CellRangeAddress range)
        {
            int row = _record.Row;
            int col = _record.Column;
            short styleIndex = _record.XFIndex;
            SetCellType(CellType.Formula, false, row, col, styleIndex);

            // Billet for formula in rec
            Ptg[] ptgsForCell = { new ExpPtg(range.FirstRow, range.FirstColumn) };
            FormulaRecordAggregate agg = (FormulaRecordAggregate)_record;
            agg.SetParsedExpression(ptgsForCell);
        }
        public CellRangeAddress ArrayFormulaRange
        {
            get
            {
                if (cellType != CellType.Formula)
                {
                    String ref1 = new CellReference(this).FormatAsString();
                    throw new InvalidOperationException("Cell " + ref1
                        + " is not part of an array formula.");
                }
                return ((FormulaRecordAggregate)_record).GetArrayFormulaRange();
            }
        }
        public ICell CopyCellTo(int targetIndex)
        {
            return this.Row.CopyCell(this.ColumnIndex,targetIndex);
        }
        /// <summary>
        /// The purpose of this method is to validate the cell state prior to modification
        /// </summary>
        /// <param name="msg"></param>
        internal void NotifyArrayFormulaChanging(String msg)
        {
            CellRangeAddress cra = this.ArrayFormulaRange;
            if (cra.NumberOfCells > 1)
            {
                throw new InvalidOperationException(msg);
            }
            //un-register the single-cell array formula from the parent XSSFSheet
            this.Row.Sheet.RemoveArrayFormula(this);
        }

        /// <summary>
        /// Called when this cell is modified.
        /// The purpose of this method is to validate the cell state prior to modification.
        /// </summary>
        internal void NotifyArrayFormulaChanging()
        {
            CellReference ref1 = new CellReference(this);
            String msg = "Cell " + ref1.FormatAsString() + " is part of a multi-cell array formula. " +
                    "You cannot change part of an array.";
            NotifyArrayFormulaChanging(msg);
        }

        public CellType GetCachedFormulaResultTypeEnum()
        {
            throw new NotImplementedException();
        }

        public bool IsMergedCell
        {
            get
            {
                foreach (CellRangeAddress range in _sheet.Sheet.MergedRecords.MergedRegions)
                {
                    if (range.FirstColumn <= this.ColumnIndex
                        && range.LastColumn >= this.ColumnIndex
                        && range.FirstRow <= this.RowIndex
                        && range.LastRow >= this.RowIndex)
                    {
                        return true;
                    }
                }
                return false;
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.Streaming.Properties;
using NPOI.XSSF.Streaming.Values;
using NPOI.XSSF.UserModel;
using Value = NPOI.XSSF.Streaming.Values.Value;

namespace NPOI.XSSF.Streaminging
{
    public class SXSSFCell : ICell
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(SXSSFCell));

        private SXSSFRow _row;
        private Value _value;
        private ICellStyle _style;
        private Property _firstProperty;

        public SXSSFCell(SXSSFRow row, CellType cellType)
        {
            _row = row;
            setType(cellType);
        }

        public CellRangeAddress ArrayFormulaRange
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool BooleanCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return false;
                    case CellType.Formula:
                        {
                            FormulaValue fv = (FormulaValue)_value;
                            if (fv.GetFormulaType() != CellType.Boolean)
                                throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.Boolean, CellType.Formula, false));
                            return ((BooleanFormulaValue)_value).PreEvaluatedValue;
                        }
                    case CellType.Boolean:
                        {
                            return ((BooleanValue)_value).Value;
                        }
                    default:
                        throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.Boolean, cellType, false));
                }
            }
        }

        private string BuildTypeMismatchMessage(CellType expectedTypeCode, CellType actualTypeCode,
            bool isFormulaCell)
        {
            return string.Format("Cannot get a {0} value from a {1} {2} cell", expectedTypeCode, actualTypeCode,(isFormulaCell ? "formula " : ""));
        }

        public CellType CachedFormulaResultType
        {
            get { return GetCachedFormulaResultTypeEnum(); }
        }

        //TODO: thismight be a little messed up.
        public CellType GetCachedFormulaResultTypeEnum()
        {
            if (_value.GetType() != CellType.Formula)
            {
                throw new InvalidOperationException("Only formula cells have cached results");
            }

            return ((FormulaValue)_value).GetFormulaType();
        }

        public IComment CellComment
        {
            get
            {
                return (IComment)GetPropertyValue(Property.COMMENT);
            }

            set
            {
                SetProperty(Property.COMMENT, value);
            }
        }

        public string CellFormula
        {
            get
            {
                if (_value.GetType() != CellType.Formula)
                    throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.Formula, _value.GetType(), false));
                return ((FormulaValue) _value).Value;
            }

            set
            {
                if (value == null)
                {
                    setType(CellType.Blank);
                    return;
                }

                EnsureFormulaType(ComputeTypeFromFormula(value));
                ((FormulaValue)_value).Value = value;
            }
        }

        public ICellStyle CellStyle
        {
            get
            {
                //TODO: seems like off functionality.
                if (_style == null)
                {
                    SXSSFWorkbook wb = (SXSSFWorkbook)Row.Sheet.Workbook;
                    return wb.GetCellStyleAt(0);
                }
                else
                {
                    return _style;
                }
            }

            set
            {
                _style = value;
            }
        }

        //TODO: can replace CellType cellType = _value.GetType();
        public CellType CellType
        {
            get { return _value.GetType(); }
        }

        public int ColumnIndex
        {
            get
            {
                return _row.getCellIndex(this);
            }
        }

        public DateTime DateCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                if (cellType == CellType.Blank)
                {
                    return new DateTime();
                }

                double value = NumericCellValue;
                bool date1904= ((XSSFWorkbook)Sheet.Workbook).IsDate1904();
                return DateUtil.GetJavaDate(value, date1904);
            }
        }

        public byte ErrorCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return 0;
                    case CellType.Formula:
                        {
                            FormulaValue fv = (FormulaValue)_value;
                            if (fv.GetFormulaType() != CellType.Error)
                                new InvalidOperationException(BuildTypeMismatchMessage(CellType.Error, CellType.Formula, false));
                            return ((ErrorFormulaValue)_value).PreEvaluatedValue;
                        }
                    case CellType.Error:
                        {
                            return ((ErrorValue)_value).Value;
                        }
                    default:
                        throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.Error, cellType, false));
                }
            }
        }

        public IHyperlink Hyperlink
        {
            get
            {
                return (IHyperlink)GetPropertyValue(Property.HYPERLINK);
            }

            set
            {
                ////TODO: may need to set private _link
                throw new NotImplementedException();
                //if (Hyperlink == null)
                //{
                //    RemoveHyperlink();
                //    return;
                //}

                //SetProperty(Property.HYPERLINK, Hyperlink);

                //XSSFHyperlink xssfobj = (XSSFHyperlink)Hyperlink;
                //// Assign to us
                //CellReference reference = new CellReference(RowIndex, ColumnIndex);
                //xssfobj.GetCTHyperlink().@ref = reference.FormatAsString();

                //// Add to the lists
                //Sheet._sh.addHyperlink(xssfobj);
            }
        }

        public bool IsMergedCell
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool IsPartOfArrayFormulaGroup
        {
            get
            {
                return false;
                throw new NotImplementedException();
            }
        }

        public double NumericCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return 0.0;
                    case CellType.Formula:
                        {
                            FormulaValue fv = (FormulaValue)_value;
                            if (fv.GetFormulaType() != CellType.Numeric)
                                throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.Numeric, CellType.Formula, false));
                            return ((NumericFormulaValue)_value).PreEvaluatedValue;
                        }
                    case CellType.Numeric:
                        return ((NumericValue) _value).Value;
                    default:
                        throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.Numeric, cellType, false));
                }
            }
        }

        public IRichTextString RichStringCellValue
        {
            get
            {
                //TODO: make sure this first line is equiv
                CellType cellType = _value.GetType();
                if (cellType != CellType.String)
                    throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.String, cellType, false));

                StringValue sval = (StringValue)_value;
                if (sval.IsRichText())
                    return ((RichTextValue) _value).Value;
                else
                {
                    string plainText = StringCellValue;
                    return Sheet.Workbook.GetCreationHelper().CreateRichTextString(plainText);
                }
            }
        }

        public IRow Row
        {
            //TODO: fixed when implemeneted
            get { return _row; }
        }

        public int RowIndex
        {
            get
            {
                return _row.RowNum;
            }
        }

        public ISheet Sheet
        {
            get
            {
                return _row.Sheet;
            }
        }

        public string StringCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return "";
                    case CellType.Formula:
                        {
                            FormulaValue fv = (FormulaValue)_value;
                            if (fv.GetFormulaType() != CellType.String)
                                throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.String, CellType.Formula, false));
                            return ((StringFormulaValue)_value).PreEvaluatedValue;
                        }
                    case CellType.String:
                    {
                        if (((StringValue) _value).IsRichText())
                            return ((RichTextValue) _value).Value.String;
                        else
                            return ((PlainStringValue) _value).Value;
                    }
                    default:
                        throw new InvalidOperationException(BuildTypeMismatchMessage(CellType.String, cellType, false));
                }

            }
        }

        public ICell CopyCellTo(int targetIndex)
        {
            throw new NotImplementedException();
        }

        public void RemoveCellComment()
        {
            RemoveProperty(Property.COMMENT);
        }

        public void RemoveHyperlink()
        {
            throw new NotImplementedException();
            //RemoveProperty(Property.HYPERLINK);

            //Sheet._sh.removeHyperlink(RowIndex, ColumnIndex);
        }

        public void SetAsActiveCell()
        {
            //TODO: Fixme!
            throw new NotImplementedException();
           // Sheet.SetActiveCell(CellAddress);
        }

        public void SetCellErrorValue(byte value)
        {
            EnsureType(CellType.Error);
            if (_value.GetType() == CellType.Formula)
                ((ErrorFormulaValue)_value).PreEvaluatedValue = value;
            else
                ((ErrorValue)_value).Value = value;
        }

        public void SetCellFormula(string formula)
        {
            if (formula == null)
            {
                setType(CellType.Blank);
                return;
            }

            EnsureFormulaType(ComputeTypeFromFormula(formula));
            ((FormulaValue)_value).Value = formula;
        }
        
        private CellType ComputeTypeFromFormula(String formula)
        {
            return CellType.Numeric;
        }
        public void SetCellType(CellType cellType)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(string value)
        {
            if (value != null)
            {
                EnsureTypeOrFormulaType(CellType.String);

                if (value.Length > SpreadsheetVersion.EXCEL2007.MaxTextLength)
                {
                    throw new InvalidOperationException("The maximum length of cell contents (text) is 32,767 characters");
                }

                if (_value.GetType() == CellType.Formula)
                    if (_value is NumericFormulaValue) {
                    ((NumericFormulaValue)_value).PreEvaluatedValue = Double.Parse(value);
                } else {
                    ((StringFormulaValue)_value).Value = value;
                }
            else
                ((PlainStringValue)_value).Value = value;
            }
            else
            {
                SetCellType(CellType.Blank);
            }
        }

        public void SetCellValue(bool value)
        {
            EnsureTypeOrFormulaType(CellType.Boolean);
            if (_value.GetType() == CellType.Formula)
                ((BooleanFormulaValue)_value).PreEvaluatedValue = value;
            else
                ((BooleanValue)_value).Value = value;
        }

        public void SetCellValue(IRichTextString value)
        {
            XSSFRichTextString xvalue = (XSSFRichTextString)value;
            
            if (xvalue != null && xvalue.String != null)
            {
                EnsureRichTextStringType();

                if (xvalue.Length > SpreadsheetVersion.EXCEL2007.MaxTextLength)
                {
                    throw new InvalidOperationException("The maximum length of cell contents (text) is 32,767 characters");
                }
                //TODO: this might just be the length > 0 investigate more.
                if (xvalue.HasFormatting())
                    logger.Log(POILogger.WARN, "SXSSF doesn't support Shared Strings, rich text formatting information has be lost");

                ((RichTextValue)_value).Value = xvalue;
            }
            else
            {
                SetCellType(CellType.Blank);
            }
        }

        public void SetCellValue(DateTime? value)
        {
            if (value == null)
            {
                SetCellType(CellType.Blank);
                return;
            }

            //TODO: add is date to IWorkbook
            bool date1904 = ((XSSFWorkbook)Sheet.Workbook).IsDate1904();
            SetCellValue(DateUtil.GetExcelDate(value.Value, date1904));
        }

        public void SetCellValue(double value)
        {
            if (Double.IsInfinity(value))
            {
                // Excel does not support positive/negative infinities,
                // rather, it gives a #DIV/0! error in these cases.
                SetCellErrorValue(FormulaError.DIV0.Code);
            }
            else if (Double.IsNaN(value))
            {
                SetCellErrorValue(FormulaError.NUM.Code);
            }
            else
            {
                EnsureTypeOrFormulaType(CellType.Numeric);
                if (_value.GetType() == CellType.Formula)
                    ((NumericFormulaValue)_value).PreEvaluatedValue = value;
                else
                    ((NumericValue)_value).Value = value;
            }
        }

        #region Weird extra stuff
        private void RemoveProperty(int type)
        {
            Property current = _firstProperty;
            Property previous = null;
            while (current != null && current.GetType() != type)
            {
                previous = current;
                current = current._next;
            }
            if (current != null)
            {
                if (previous != null)
                {
                    previous._next = current._next;
                }
                else
                {
                    _firstProperty = current._next;
                }
            }
        }

        //TODO; type needs to be an enum or the class type
        private void SetProperty(int type, object value)
        {
            Property current = _firstProperty;
            Property previous = null;
            while (current != null && current.GetType() != type)
            {
                previous = current;
                current = current._next;
            }
            if (current != null)
            {
                current._value = value;
            }
            else
            {
                switch (type)
                {
                    case Property.COMMENT:
                        {
                            current = new CommentProperty(value);
                            break;
                        }
                    case Property.HYPERLINK:
                        {
                            current = new HyperlinkProperty(value);
                            break;
                        }
                    default:
                        {
                            throw new InvalidOperationException("Invalid type: " + type);
                        }
                }
                if (previous != null)
                {
                    previous._next = current;
                }
                else
                {
                    _firstProperty = current;
                }
            }
        }

        private object GetPropertyValue(int type)
        {
            return getPropertyValue(type, null);
        }

        private object getPropertyValue(int type, string defaultValue)
        {
            Property current = _firstProperty;
            while (current != null && current.GetType() != type) current = current._next;
            return current == null ? defaultValue : current._value;
        }

        private void EnsurePlainStringType()
        {
            if (_value.GetType() != CellType.String
               || ((StringValue)_value).IsRichText())
                _value = new PlainStringValue();
        }

        private void EnsureRichTextStringType()
        {
            if (_value.GetType() != CellType.String
               || !((StringValue)_value).IsRichText())
                _value = new RichTextValue();
        }

        private void EnsureType(CellType type)
        {
            if (_value.GetType() != type)
                setType(type);
        }

        private void EnsureFormulaType(CellType type)
        {
            if (_value.GetType() != CellType.Formula
               || ((FormulaValue)_value).GetFormulaType() != type)
                setFormulaType(type);
        }
        /*
         * Sets the cell type to type if it is different
         */

        private void EnsureTypeOrFormulaType(CellType type)
        {
            if (_value.GetType() == type)
            {
                if (type == CellType.String && ((StringValue)_value).IsRichText())
                    setType(CellType.String);
                return;
            }
            if (_value.GetType() == CellType.Formula)
            {
                if (((FormulaValue)_value).GetFormulaType() == type)
                    return;
                setFormulaType(type); // once a formula, always a formula
                return;
            }
            setType(type);
        }

        /*package*/
        private void setType(CellType type)
        {
            switch (type)
            {
                case CellType.Numeric:
                    {
                        _value = new NumericValue();
                        break;
                    }
                case CellType.String:
                    {
                        PlainStringValue sval = new PlainStringValue();
                        if (_value != null)
                        {
                            // if a cell is not blank then convert the old value to string
                            String str = convertCellValueToString();
                            sval.Value = str;
                        }
                        _value = sval;
                        break;
                    }
                case CellType.Formula:
                    {
                        _value = new NumericFormulaValue();
                        break;
                    }
                case CellType.Blank:
                    {
                        _value = new BlankValue();
                        break;
                    }
                case CellType.Boolean:
                    {
                        BooleanValue bval = new BooleanValue();
                        if (_value != null)
                        {
                            // if a cell is not blank then convert the old value to string
                            bool val = convertCellValueToBoolean();
                            bval.Value = val;
                        }
                        _value = bval;
                        break;
                    }
                case CellType.Error:
                    {
                        _value = new ErrorValue();
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException("Illegal type " + type);
                    }
            }
        }

        private void setFormulaType(CellType type)
        {
            Value prevValue = _value;
            switch (type)
            {
                case CellType.Numeric:
                    {
                        _value = new NumericFormulaValue();
                        break;
                    }
                case CellType.String:
                    {
                        _value = new StringFormulaValue();
                        break;
                    }
                case CellType.Boolean:
                    {
                        _value = new BooleanFormulaValue();
                        break;
                    }
                case CellType.Error:
                    {
                        _value = new ErrorFormulaValue();
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException("Illegal type " + type);
                    }
            }

            // if we had a Formula before, we should copy over the _value of the formula
            if (prevValue is FormulaValue)
            {
                ((FormulaValue)_value).Value = ((FormulaValue)prevValue).Value;
            }
        }

        //COPIED FROM https://svn.apache.org/repos/asf/poi/trunk/src/ooxml/java/org/apache/poi/xssf/usermodel/XSSFCell.java since the functions are declared private there
        //TODO: may already exist in NPOI

        private bool convertCellValueToBoolean()
        {
            CellType cellType = _value.GetType();

            if (cellType == CellType.Formula)
            {
                cellType = GetCachedFormulaResultTypeEnum();
            }

            switch (cellType)
            {
                case CellType.Boolean:
                    return BooleanCellValue;
                case CellType.String:

                    String text = StringCellValue;
                    return Boolean.Parse(text);
                case CellType.Numeric:
                    return NumericCellValue != 0;
                case CellType.Error:
                case CellType.Blank:
                    return false;
                default: throw new RuntimeException("Unexpected cell type (" + cellType + ")");
            }

        }
        private String convertCellValueToString()
        {
            CellType cellType = _value.GetType();
            return convertCellValueToString(cellType);
        }
        private String convertCellValueToString(CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.String:
                    return StringCellValue;
                case CellType.Numeric:
                    return NumericCellValue.ToString();
                case CellType.Error:
                    byte errVal = ErrorCellValue;
                    return FormulaError.ForInt(errVal).ToString();

                case CellType.Formula:
                    if (_value != null)
                    {
                        FormulaValue fv = (FormulaValue)_value;
                        if (fv.GetFormulaType() != CellType.Formula)
                        {
                            return convertCellValueToString(fv.GetFormulaType());
                        }
                    }
                    return "";
                default:
                    throw new InvalidOperationException("Unexpected cell type (" + cellType + ")");
            }
        }

        //END OF COPIED CODE
        #endregion


        public void SetCellValue(DateTime value)
        {
            throw new NotImplementedException();
        }
    }
}

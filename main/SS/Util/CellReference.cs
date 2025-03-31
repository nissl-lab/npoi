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



namespace NPOI.SS.Util
{
    using System;
    using System.Text; 
using Cysharp.Text;
    using System.Text.RegularExpressions;
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using System.Globalization;

    public enum NameType:int
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        None = 0,
        Cell = 1,
        NamedRange = 2,
        Column = 3,
        Row = 4,
        BadCellOrNamedRange = -1
    }

    /**
     * <p>Common conversion functions between Excel style A1, C27 style
     *  cell references, and POI usermodel style row=0, column=0
     *  style references. Handles sheet-based and sheet-free references
     *  as well, eg "Sheet1!A1" and "$B$72"</p>
     *  
     *  <p>Use <tt>CellReference</tt> when the concept of
     * relative/absolute does apply (such as a cell reference in a formula).
     * Use {@link CellAddress} when you want to refer to the location of a cell in a sheet
     * when the concept of relative/absolute does not apply (such as the anchor location 
     * of a cell comment).
     * <tt>CellReference</tt>s have a concept of "sheet", while <tt>CellAddress</tt>es do not.</p>
     */

    public class CellReference
    {
        /** The character ($) that signifies a row or column value is absolute instead of relative */
        private const char ABSOLUTE_REFERENCE_MARKER = '$';
        /** The character (!) that Separates sheet names from cell references */
        private const char SHEET_NAME_DELIMITER = '!';
        /** The character (') used to quote sheet names when they contain special characters */
        private const char SPECIAL_NAME_DELIMITER = '\'';

        //private static final Pattern NAMED_RANGE_NAME_PATTERN = Pattern.compile("[_A-Z][_.A-Z0-9]*", Pattern.CASE_INSENSITIVE);
        private static readonly Regex NAMED_RANGE_NAME_PATTERN = new Regex("^[_A-Za-z][_.A-Za-z0-9]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);
        //private static string BIFF8_LAST_COLUMN = "IV";
        //private static int BIFF8_LAST_COLUMN_TEXT_LEN = BIFF8_LAST_COLUMN.Length;
        //private static string BIFF8_LAST_ROW = (0x10000).ToString();
        //private static int BIFF8_LAST_ROW_TEXT_LEN = BIFF8_LAST_ROW.Length;

        // FIXME: _sheetName may be null, depending on the entry point.
        // Perhaps it would be better to declare _sheetName is never null, using an empty string to represent a 2D reference.
        private readonly String _sheetName;
        private readonly int _rowIndex;
        private readonly int _colIndex;
        private readonly bool _isRowAbs;
        private readonly bool _isColAbs;

        /**
         * Create an cell ref from a string representation.  Sheet names containing special characters should be
         * delimited and escaped as per normal syntax rules for formulas.
         */
        public CellReference(String cellRef) : this(cellRef.AsSpan())
        {
        }

        public CellReference(ReadOnlySpan<char> cellRef)
        {
            if (cellRef.EndsWith("#REF!".AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Cell reference invalid: " + cellRef.ToString());
            }
            CellRefPartsInner parts = SeparateRefParts(cellRef);
            _sheetName = parts.sheetName;//parts[0];
            ReadOnlySpan<char> colRef = parts.colRef;// parts[1];
            _isColAbs = parts.columnPrefix == '$';
            if (colRef.Length == 0)
            {
                _colIndex = -1;
            }
            else
            {
                _colIndex = ConvertColStringToIndex(colRef);
            }


            ReadOnlySpan<char> rowRef = parts.rowRef;// parts[2];
            _isRowAbs = parts.rowPrefix == '$';
            if (rowRef.Length == 0)
            {
                _rowIndex = -1;
            }
            else
            {
                CellReferenceParser.TryParsePositiveInt32Fast(rowRef, out var rowRefNumber);
                _rowIndex = rowRefNumber - 1; // -1 to convert 1-based to zero-based
            }
        }
        public CellReference(ICell cell):this(cell.RowIndex, cell.ColumnIndex, false, false)
        {
            
        }
        public CellReference(int pRow, int pCol)
            : this(pRow, pCol, false, false)
        {

        }
        public CellReference(int pRow, short pCol)
            : this(pRow, pCol & 0xFFFF, false, false)
        {

        }
        public CellReference(int pRow, int pCol, bool pAbsRow, bool pAbsCol)
            : this(null, pRow, pCol, pAbsRow, pAbsCol)
        {

        }
        public CellReference(String pSheetName, int pRow, int pCol, bool pAbsRow, bool pAbsCol)
        {
            // TODO - "-1" is a special value being temporarily used for whole row and whole column area references.
            // so these Checks are currently N.Q.R.
            if (pRow < -1)
            {
                throw new ArgumentException("row index may not be negative, but had " + pRow);
            }
            if (pCol < -1)
            {
                throw new ArgumentException("column index may not be negative, but had " + pCol);
            }
            _sheetName = pSheetName;
            _rowIndex = pRow;
            _colIndex = pCol;
            _isRowAbs = pAbsRow;
            _isColAbs = pAbsCol;
        }

        public int Row
        {
            get { return _rowIndex; }
        }
        public short Col
        {
            get
            {
                return (short)_colIndex;
            }
        }
        public bool IsRowAbsolute
        {
            get { return _isRowAbs; }
        }
        public bool IsColAbsolute
        {
            get { return _isColAbs; }
        }
        /**
          * @return possibly <c>null</c> if this is a 2D reference.  Special characters are not
          * escaped or delimited
          */
        public String SheetName
        {
            get { return _sheetName; }
        }

        /**
         * takes in a column reference portion of a CellRef and converts it from
         * ALPHA-26 number format to 0-based base 10.
         * 'A' -> 0
         * 'Z' -> 25
         * 'AA' -> 26
         * 'IV' -> 255
         * @return zero based column index
         */
        public static int ConvertColStringToIndex(String refs) => ConvertColStringToIndex(refs.AsSpan());

        public static int ConvertColStringToIndex(ReadOnlySpan<char> refs)
        {
            int retval = 0;
            for (int k = 0; k < refs.Length; k++)
            {
                char thechar = char.ToUpperInvariant(refs[k]);
                if (thechar == ABSOLUTE_REFERENCE_MARKER)
                {
                    if (k != 0)
                    {
                        throw new ArgumentException("Bad col ref format '" + refs.ToString() + "'");
                    }
                    continue;
                }

                // Character is uppercase letter, find relative value to A
                retval = (retval * 26) + (thechar - 'A' + 1);
            }
            return retval - 1;
        }

        public static bool IsPartAbsolute(String part)
        {
            return part[0] == ABSOLUTE_REFERENCE_MARKER;
        }

        public static NameType ClassifyCellReference(String str, SpreadsheetVersion ssVersion) => ClassifyCellReference(str.AsSpan(), ssVersion);

        public static NameType ClassifyCellReference(ReadOnlySpan<char> str, SpreadsheetVersion ssVersion)
        {
            int len = str.Length;
            if (len < 1)
            {
                throw new ArgumentException("Empty string not allowed");
            }
            char firstChar = str[0];
            switch (firstChar)
            {
                case ABSOLUTE_REFERENCE_MARKER:
                case '.':
                case '_':
                    break;
                default:
                    if (!Char.IsLetter(firstChar) && !Char.IsDigit(firstChar))
                    {
                        throw new ArgumentException("Invalid first char (" + firstChar
                                + ") of cell reference or named range.  Letter expected");
                    }
                    break;
            }
            if (!Char.IsDigit(str[len - 1]))
            {
                // no digits at end of str
                return ValidateNamedRangeName(str, ssVersion);
            }

            if (!CellReferenceParser.TryParseStrictCellReference(str, out var lettersGroup, out var digitsGroup))
            {
                return ValidateNamedRangeName(str, ssVersion);
            }

            if (CellReferenceIsWithinRange(lettersGroup, digitsGroup, ssVersion))
            {
                // valid cell reference
                return NameType.Cell;
            }
            // If str looks like a cell reference, but is out of (row/col) range, it is a valid
            // named range name
            // This behaviour is a little weird.  For example, "IW123" is a valid named range name
            // because the column "IW" is beyond the maximum "IV".  Note - this behaviour is version
            // dependent.  In BIFF12, "IW123" is not a valid named range name, but in BIFF8 it is.
            if (str.IndexOf(ABSOLUTE_REFERENCE_MARKER) >= 0)
            {
                // Of course, named range names cannot have '$'
                return NameType.BadCellOrNamedRange;
            }
            return NameType.NamedRange;
        }

        private static NameType ValidateNamedRangeName(ReadOnlySpan<char> str, SpreadsheetVersion ssVersion)
        {
            if (CellReferenceParser.TryParseColumnReference(str, out var colStr))
            {
                if (IsColumnWithinRange(colStr, ssVersion))
                {
                    return NameType.Column;
                }
            }
            if (CellReferenceParser.TryParseRowReference(str, out var rowStr))
            {
                if (IsRowWithinRange(rowStr, ssVersion))
                {
                    return NameType.Row;
                }
            }
            // TODO
            if (!NAMED_RANGE_NAME_PATTERN.IsMatch(str.ToString()))
            {
                return NameType.BadCellOrNamedRange;
            }
            return NameType.NamedRange;
        }
        /**
         * Takes in a 0-based base-10 column and returns a ALPHA-26
         *  representation.
         * eg column #3 -> D
         */
        public static String ConvertNumToColString(int col)
        {
            // Excel counts column A as the 1st column, we
            //  treat it as the 0th one
            int excelColNum = col + 1;

            StringBuilder colRef = new StringBuilder(2);
            int colRemain = excelColNum;

            while (colRemain > 0)
            {
                int thisPart = colRemain % 26;
                if (thisPart == 0) { thisPart = 26; }
                colRemain = (colRemain - thisPart) / 26;

                // The letter A is at 65
                char colChar = (char)(thisPart + 64);
                colRef.Insert(0, colChar);
            }

            return colRef.ToString();
        }
        internal readonly ref struct CellRefPartsInner
        {
            public readonly String sheetName;
            public readonly char rowPrefix;
            public readonly char columnPrefix;
            public readonly ReadOnlySpan<char> rowRef;
            public readonly ReadOnlySpan<char> colRef;

            public CellRefPartsInner(string sheetName, char rowPrefix, ReadOnlySpan<char> rowRef, char columnPrefix, ReadOnlySpan<char> colRef)
            {
                this.sheetName = sheetName;
                this.rowPrefix = rowPrefix;
                this.rowRef = rowRef;
                this.columnPrefix = columnPrefix;
                this.colRef = colRef;
            }
        }
        /**
         * Separates the row from the columns and returns an array of three Strings.  The first element
         * is the sheet name. Only the first element may be null.  The second element in is the column 
         * name still in ALPHA-26 number format.  The third element is the row.
         */
        private static CellRefPartsInner SeparateRefParts(ReadOnlySpan<char> reference)
        {

            int plingPos = reference.LastIndexOf(SHEET_NAME_DELIMITER);
            String sheetName = ParseSheetName(reference, plingPos);
            int start = plingPos + 1;
            String cell = reference.ToString().Substring(plingPos + 1).ToUpper(CultureInfo.InvariantCulture);
            if (!CellReferenceParser.TryParseCellReference(cell.AsSpan(), out var columnPrefix, out var column, out var rowPrefix, out var row))
                throw new ArgumentException("Invalid CellReference: " + reference.ToString());

            CellRefPartsInner cellRefParts = new CellRefPartsInner(sheetName, rowPrefix, row, columnPrefix, column);
            return cellRefParts;
        }

        private static String ParseSheetName(ReadOnlySpan<char> reference, int indexOfSheetNameDelimiter)
        {
            if (indexOfSheetNameDelimiter < 0)
            {
                return null;
            }

            bool IsQuoted = reference[0] == SPECIAL_NAME_DELIMITER;
            if (!IsQuoted)
            {
                // sheet names with spaces must be quoted
                if (reference.IndexOf(' ') == -1)
                {
                    return reference.Slice(0, indexOfSheetNameDelimiter).ToString();
                }
                else
                {
                    throw new ArgumentException("Sheet names containing spaces must be quoted: (" + reference.ToString() + ")");
                }
            }
            int lastQuotePos = indexOfSheetNameDelimiter - 1;
            if (reference[lastQuotePos] != SPECIAL_NAME_DELIMITER)
            {
                throw new ArgumentException("Mismatched quotes: (" + reference.ToString() + ")");
            }

            // TODO - refactor cell reference parsing logic to one place.
            // Current known incarnations: 
            //   FormulaParser.Name
            //   CellReference.ParseSheetName() (here)
            //   AreaReference.SeparateAreaRefs() 
            //   SheetNameFormatter.format() (inverse)

            StringBuilder sb = new StringBuilder(indexOfSheetNameDelimiter);

            for (int i = 1; i < lastQuotePos; i++)
            { // Note boundaries - skip outer quotes
                char ch = reference[i];
                if (ch != SPECIAL_NAME_DELIMITER)
                {
                    sb.Append(ch);
                    continue;
                }
                if (i < lastQuotePos)
                {
                    if (reference[i + 1] == SPECIAL_NAME_DELIMITER)
                    {
                        // two consecutive quotes is the escape sequence for a single one
                        i++; // skip this and keep parsing the special name
                        sb.Append(ch);
                        continue;
                    }
                }
                throw new ArgumentException("Bad sheet name quote escaping: (" + reference.ToString() + ")");
            }
            return sb.ToString();
        }


        /**
         *  Example return values:
         *    <table border="0" cellpAdding="1" cellspacing="0" summary="Example return values">
         *      <tr><th align='left'>Result</th><th align='left'>Comment</th></tr>
         *      <tr><td>A1</td><td>Cell reference without sheet</td></tr>
         *      <tr><td>Sheet1!A1</td><td>Standard sheet name</td></tr>
         *      <tr><td>'O''Brien''s Sales'!A1'</td><td>Sheet name with special characters</td></tr>
         *    </table>
         * @return the text representation of this cell reference as it would appear in a formula.
         */
        public String FormatAsString()
        {

            StringBuilder sb = new StringBuilder(32);
            if (_sheetName != null)
            {
                SheetNameFormatter.AppendFormat(sb, _sheetName);
                sb.Append(SHEET_NAME_DELIMITER);
            }
            AppendCellReference(sb);
            return sb.ToString();
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(this.GetType().Name).Append(" [");
            sb.Append(FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }

        /**
         * Returns the three parts of the cell reference, the
         *  Sheet name (or null if none supplied), the 1 based
         *  row number, and the A based column letter.
         * This will not include any markers for absolute
         *  references, so use {@link #formatAsString()}
         *  to properly turn references into strings. 
         */
        public String[] CellRefParts
        {
            get
            {
                return new String[] {
                    _sheetName,
                    (_rowIndex+1).ToString(CultureInfo.InvariantCulture),
                   ConvertNumToColString(_colIndex)
                };
            }
        }

        /**
         * Appends cell reference with '$' markers for absolute values as required.
         * Sheet name is not included.
         */
        /* package */
        public void AppendCellReference(StringBuilder sb)
        {
            if (_colIndex != -1)
            {
                if (_isColAbs)
                {
                    sb.Append(ABSOLUTE_REFERENCE_MARKER);
                }
                sb.Append(ConvertNumToColString(_colIndex));
            }

            if (_rowIndex != -1)
            {
                if (_isRowAbs)
                {
                    sb.Append(ABSOLUTE_REFERENCE_MARKER);
                }
                sb.Append(_rowIndex + 1);
            }
        }

        /**
         * Used to decide whether a name of the form "[A-Z]*[0-9]*" that appears in a formula can be 
         * interpreted as a cell reference.  Names of that form can be also used for sheets and/or
         * named ranges, and in those circumstances, the question of whether the potential cell 
         * reference is valid (in range) becomes important.
         * <p/>
         * Note - that the maximum sheet size varies across Excel versions:
         * <p/>
         * <blockquote><table border="0" cellpadding="1" cellspacing="0" 
         *                 summary="Notable cases.">
         *   <tr><th>Version </th><th>File Format </th>
         *   	<th>Last Column </th><th>Last Row</th></tr>
         *   <tr><td>97-2003</td><td>BIFF8</td><td>"IV" (2^8)</td><td>65536 (2^14)</td></tr>
         *   <tr><td>2007</td><td>BIFF12</td><td>"XFD" (2^14)</td><td>1048576 (2^20)</td></tr>
         * </table></blockquote>
         * POI currently targets BIFF8 (Excel 97-2003), so the following behaviour can be observed for
         * this method:
         * <blockquote><table border="0" cellpadding="1" cellspacing="0" 
         *                 summary="Notable cases.">
         *   <tr><th>Input    </th>
         *       <th>Result </th></tr>
         *   <tr><td>"A", "1"</td><td>true</td></tr>
         *   <tr><td>"a", "111"</td><td>true</td></tr>
         *   <tr><td>"A", "65536"</td><td>true</td></tr>
         *   <tr><td>"A", "65537"</td><td>false</td></tr>
         *   <tr><td>"iv", "1"</td><td>true</td></tr>
         *   <tr><td>"IW", "1"</td><td>false</td></tr>
         *   <tr><td>"AAA", "1"</td><td>false</td></tr>
         *   <tr><td>"a", "111"</td><td>true</td></tr>
         *   <tr><td>"Sheet", "1"</td><td>false</td></tr>
         * </table></blockquote>
         * 
         * @param colStr a string of only letter characters
         * @param rowStr a string of only digit characters
         * @return <c>true</c> if the row and col parameters are within range of a BIFF8 spreadsheet.
         */
        public static bool CellReferenceIsWithinRange(String colStr, String rowStr, SpreadsheetVersion ssVersion)
            => CellReferenceIsWithinRange(colStr.AsSpan(), rowStr.AsSpan(), ssVersion);

        public static bool CellReferenceIsWithinRange(ReadOnlySpan<char> colStr, ReadOnlySpan<char> rowStr, SpreadsheetVersion ssVersion)
        {
            if (!IsColumnWithinRange(colStr, ssVersion))
            {
                return false;
            }
            return IsRowWithinRange(rowStr, ssVersion);
        }

        /**
         * @deprecated 3.15 beta 2. Use {@link #isColumnWithinRange}.
         */
        [Obsolete("deprecated 3.15 beta 2. Use {@link #isColumnWithinRange}.")]
        public static bool IsColumnWithnRange(String colStr, SpreadsheetVersion ssVersion)
        {
            return IsColumnWithinRange(colStr, ssVersion);
        }

        public static bool IsRowWithinRange(String rowStr, SpreadsheetVersion ssVersion)
            => IsRowWithinRange(rowStr.AsSpan(), ssVersion);

        public static bool IsRowWithinRange(ReadOnlySpan<char> rowStr, SpreadsheetVersion ssVersion)
        {
            CellReferenceParser.TryParsePositiveInt32Fast(rowStr, out var rowNum);
            rowNum -= 1;
            return 0 <= rowNum && rowNum <= ssVersion.LastRowIndex;
        }

        [Obsolete("deprecated 3.15 beta 2. Use {@link #isRowWithinRange}")]
        public static bool isRowWithnRange(String rowStr, SpreadsheetVersion ssVersion)
        {
            return IsRowWithinRange(rowStr, ssVersion);
        }

        public static bool IsColumnWithinRange(String colStr, SpreadsheetVersion ssVersion)
            => IsColumnWithinRange(colStr.AsSpan(), ssVersion);

        public static bool IsColumnWithinRange(ReadOnlySpan<char> colStr, SpreadsheetVersion ssVersion)
        {
            String lastCol = ssVersion.LastColumnName;
            int lastColLength = lastCol.Length;

            int numberOfLetters = colStr.Length;
            if (numberOfLetters > lastColLength)
            {
                // "Sheet1" case etc
                return false; // that was easy
            }
            if (numberOfLetters == lastColLength)
            {
                //if (colStr.ToUpper().CompareTo(lastCol) > 0)
                if (colStr.CompareTo(lastCol.AsSpan(), StringComparison.OrdinalIgnoreCase) > 0)
                {
                    return false;
                }
            }
            else
            {
                // apparent column name has less chars than max
                // no need to check range
            }
            return true;
        }
        public override bool Equals(Object o)
        {
            if (object.ReferenceEquals(this, o))
                return true;
            if (o is not CellReference cr)
            {
                return false;
            }

            return _rowIndex == cr._rowIndex
                   && _colIndex == cr._colIndex
                   && _isRowAbs == cr._isRowAbs
                   && _isColAbs == cr._isColAbs
                   && ((_sheetName == null)
                       ? (cr._sheetName == null)
                       : _sheetName.Equals(cr._sheetName));
        }

        public override int GetHashCode ()
        {
            int result = 17;
            result = 31 * result + _rowIndex;
            result = 31 * result + _colIndex;
            result = 31 * result + (_isRowAbs ? 1 : 0);
            result = 31 * result + (_isColAbs ? 1 : 0);
            result = 31 * result + (_sheetName == null ? 0 : _sheetName.GetHashCode());
            return result;
        }
    }
}
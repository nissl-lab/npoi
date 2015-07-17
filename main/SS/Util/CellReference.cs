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
     *
     * @author  Avik Sengupta
     * @author  Dennis doubleday (patch to seperateRowColumns())
     */
    public class CellReference
    {
        /** The character ($) that signifies a row or column value is absolute instead of relative */
        private const char ABSOLUTE_REFERENCE_MARKER = '$';
        /** The character (!) that Separates sheet names from cell references */
        private const char SHEET_NAME_DELIMITER = '!';
        /** The character (') used to quote sheet names when they contain special characters */
        private const char SPECIAL_NAME_DELIMITER = '\'';

        /**
         * Matches a run of one or more letters followed by a run of one or more digits.
         * The run of letters is group 1 and the run of digits is group 2.  
         * Each group may optionally be prefixed with a single '$'.
         */
        private const string CELL_REF_PATTERN = @"^\$?([A-Za-z]+)\$?([0-9]+)";
        /**
         * Matches a run of one or more letters.  The run of letters is group 1.  
         * The text may optionally be prefixed with a single '$'.
         */
        private const string COLUMN_REF_PATTERN = @"^\$?([A-Za-z]+)$";
        /**
 * Matches a run of one or more digits.  The run of digits is group 1.
 * The text may optionally be prefixed with a single '$'.
 */
        private const string ROW_REF_PATTERN = @"^\$?([0-9]+)$";
        /**
         * Named range names must start with a letter or underscore.  Subsequent characters may include
         * digits or dot.  (They can even end in dot).
         */
        private const string NAMED_RANGE_NAME_PATTERN = "^[_A-Za-z][_.A-Za-z0-9]*$";
        //private static string BIFF8_LAST_COLUMN = "IV";
        //private static int BIFF8_LAST_COLUMN_TEXT_LEN = BIFF8_LAST_COLUMN.Length;
        //private static string BIFF8_LAST_ROW = (0x10000).ToString();
        //private static int BIFF8_LAST_ROW_TEXT_LEN = BIFF8_LAST_ROW.Length;


        private int _rowIndex;
        private int _colIndex;
        private String _sheetName;
        private bool _isRowAbs;
        private bool _isColAbs;

        /**
         * Create an cell ref from a string representation.  Sheet names containing special characters should be
         * delimited and escaped as per normal syntax rules for formulas.
         */
        public CellReference(String cellRef)
        {
            if (cellRef.EndsWith("#REF!", StringComparison.CurrentCulture))
            {
                throw new ArgumentException("Cell reference invalid: " + cellRef);
            }
            String[] parts = SeparateRefParts(cellRef);
            _sheetName = parts[0];
            String colRef = parts[1];
            //if (colRef.Length < 1)
            //{
            //    throw new ArgumentException("Invalid Formula cell reference: '" + cellRef + "'");
            //}
            _isColAbs = (colRef.Length > 0) && colRef[0] == '$';
            //_isColAbs = colRef[0] == '$';
            if (_isColAbs)
            {
                colRef = colRef.Substring(1);
            }
            if (colRef.Length == 0)
            {
                _colIndex = -1;
            }
            else
            {
                _colIndex = ConvertColStringToIndex(colRef);
            }
            

            String rowRef = parts[2];
            //if (rowRef.Length < 1)
            //{
            //    throw new ArgumentException("Invalid Formula cell reference: '" + cellRef + "'");
            //}
            //_isRowAbs = rowRef[0] == '$';
            _isRowAbs = (rowRef.Length > 0) && rowRef[0] == '$';
            if (_isRowAbs)
            {
                rowRef = rowRef.Substring(1);
            }
            if (rowRef.Length == 0)
            {
                _rowIndex = -1;
            }
            else
            {
                _rowIndex = int.Parse(rowRef, CultureInfo.InvariantCulture) - 1; // -1 to convert 1-based to zero-based
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
        public static int ConvertColStringToIndex(String ref1)
        {
            int retval = 0;
            char[] refs = ref1.ToUpper().ToCharArray();
            for (int k = 0; k < refs.Length; k++)
            {
                char thechar = refs[k];
                if (thechar == ABSOLUTE_REFERENCE_MARKER)
                {
                    if (k != 0)
                    {
                        throw new ArgumentException("Bad col ref format '" + ref1 + "'");
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
        public static NameType ClassifyCellReference(String str, SpreadsheetVersion ssVersion)
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
            Regex cellRefPatternMatcher = new Regex(CELL_REF_PATTERN);
            if (!cellRefPatternMatcher.IsMatch(str))
            {
                return ValidateNamedRangeName(str, ssVersion);
            }
            MatchCollection matches = cellRefPatternMatcher.Matches(str);
            string lettersGroup = matches[0].Groups[1].Value;
            string digitsGroup = matches[0].Groups[2].Value;
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
        private static NameType ValidateNamedRangeName(String str, SpreadsheetVersion ssVersion)
        {
            Regex colMatcher = new Regex(COLUMN_REF_PATTERN);

            if (colMatcher.IsMatch(str))
            {
                Group colStr = colMatcher.Matches(str)[0].Groups[1];
                if (IsColumnWithnRange(colStr.Value, ssVersion))
                {
                    return NameType.Column;
                }
            }
            Regex rowMatcher = new Regex(ROW_REF_PATTERN);
            if (rowMatcher.IsMatch(str))
            {
                Group rowStr = rowMatcher.Matches(str)[0].Groups[1];
                if (IsRowWithnRange(rowStr.Value, ssVersion))
                {
                    return NameType.Row;
                }
            }
            if (!Regex.IsMatch(str, NAMED_RANGE_NAME_PATTERN))
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

        /**
         * Separates the row from the columns and returns an array of three Strings.  The first element
         * is the sheet name. Only the first element may be null.  The second element in is the column 
         * name still in ALPHA-26 number format.  The third element is the row.
         */
        private static String[] SeparateRefParts(String reference)
        {

            int plingPos = reference.LastIndexOf(SHEET_NAME_DELIMITER);
            String sheetName = ParseSheetName(reference, plingPos);
            int start = plingPos + 1;

            int Length = reference.Length;


            int loc = start;
            // skip initial dollars 
            if (reference[loc] == ABSOLUTE_REFERENCE_MARKER)
            {
                loc++;
            }
            // step over column name chars Until first digit (or dollars) for row number.
            for (; loc < Length; loc++)
            {
                char ch = reference[loc];
                if (Char.IsDigit(ch) || ch == ABSOLUTE_REFERENCE_MARKER)
                {
                    break;
                }
            }
            return new String[] {
               sheetName,
               reference.Substring(start,loc-start),
               reference.Substring(loc),
            };
        }

        private static String ParseSheetName(String reference, int indexOfSheetNameDelimiter)
        {
            if (indexOfSheetNameDelimiter < 0)
            {
                return null;
            }

            bool IsQuoted = reference[0] == SPECIAL_NAME_DELIMITER;
            if (!IsQuoted)
            {
                return reference.Substring(0, indexOfSheetNameDelimiter);
            }
            int lastQuotePos = indexOfSheetNameDelimiter - 1;
            if (reference[lastQuotePos] != SPECIAL_NAME_DELIMITER)
            {
                throw new Exception("Mismatched quotes: (" + reference + ")");
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
                throw new Exception("Bad sheet name quote escaping: (" + reference + ")");
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
        {
            if (!IsColumnWithnRange(colStr, ssVersion))
            {
                return false;
            }
            return IsRowWithnRange(rowStr, ssVersion);
        }
        public static bool IsRowWithnRange(String rowStr, SpreadsheetVersion ssVersion)
        {
            int rowNum = Int32.Parse(rowStr, CultureInfo.InvariantCulture);

            if (rowNum < 0)
            {
                throw new InvalidOperationException("Invalid rowStr '" + rowStr + "'.");
            }
            if (rowNum == 0)
            {
                // execution Gets here because caller does first pass of discriminating
                // potential cell references using a simplistic regex pattern.
                return false;
            }
            return rowNum <= ssVersion.MaxRows;
        }

        public static bool IsColumnWithnRange(String colStr, SpreadsheetVersion ssVersion)
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
                if (string.Compare(colStr.ToUpper(), lastCol, StringComparison.Ordinal) > 0)
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
            if (!(o is CellReference))
            {
                return false;
            }
            CellReference cr = (CellReference)o;
            return _rowIndex == cr._rowIndex
                && _colIndex == cr._colIndex
                && _isRowAbs == cr._isRowAbs
                && _isColAbs == cr._isColAbs;
        }

        public override int GetHashCode ()
        {
            int result = 17;
            result = 31 * result + _rowIndex;
            result = 31 * result + _colIndex;
            result = 31 * result + (_isRowAbs ? 1 : 0);
            result = 31 * result + (_isColAbs ? 1 : 0);
            return result;
        }
    }
}
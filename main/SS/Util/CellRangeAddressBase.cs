namespace NPOI.SS.Util
{
    using System;
    /**
     * See OOO documentation: excelfileformat.pdf sec 2.5.14 - 'Cell Range Address'<p/>
     * 
     * Common subclass of 8-bit and 16-bit versions
     * 
     * @author Josh Micich
     */
    public abstract class CellRangeAddressBase
    {

        // /** max 65536 rows in BIFF8 */
        //private const int LAST_ROW_INDEX = 0x00FFFF;
        // /** max 256 columns in BIFF8 */
        //private const int LAST_COLUMN_INDEX = 0x00FF;

        private int _firstRow;
        private int _firstCol;
        private int _lastRow;
        private int _lastCol;

        protected CellRangeAddressBase(int firstRow, int lastRow, int firstCol, int lastCol)
        {
            //if (!IsValid(firstRow, lastRow, firstCol, lastCol))
            //{
            //    throw new ArgumentException("invalid cell range (" + firstRow + ", " + lastRow
            //            + ", " + firstCol + ", " + lastCol + ")");
            //}
            _firstRow = firstRow;
            _lastRow = lastRow;
            _firstCol = firstCol;
            _lastCol = lastCol;
        }
        //private static bool IsValid(int firstRow, int lastRow, int firstColumn, int lastColumn)
        //{
        //    if (lastRow < 0 || lastRow > LAST_ROW_INDEX)
        //    {
        //        return false;
        //    }
        //    if (firstRow < 0 || firstRow > LAST_ROW_INDEX)
        //    {
        //        return false;
        //    }

        //    if (lastColumn < 0 || lastColumn > LAST_COLUMN_INDEX)
        //    {
        //        return false;
        //    }
        //    if (firstColumn < 0 || firstColumn > LAST_COLUMN_INDEX)
        //    {
        //        return false;
        //    }
        //    return true;
        //}

        /**
	 * Validate the range limits against the supplied version of Excel
	 *
	 * @param ssVersion the version of Excel to validate against
	 * @throws IllegalArgumentException if the range limits are outside of the allowed range
	 */
        public void Validate(SpreadsheetVersion ssVersion)
        {
            ValidateRow(_firstRow, ssVersion);
            ValidateRow(_lastRow, ssVersion);
            ValidateColumn(_firstCol, ssVersion);
            ValidateColumn(_lastCol, ssVersion);
        }
        /**
	 * Runs a bounds check for row numbers
	 * @param row
	 */
        private static void ValidateRow(int row, SpreadsheetVersion ssVersion)
        {
            int maxrow = ssVersion.LastRowIndex;
            if (row > maxrow) throw new ArgumentException("Maximum row number is " + maxrow);
            if (row < 0) throw new ArgumentException("Minumum row number is 0");
        }

        /**
         * Runs a bounds check for column numbers
         * @param column
         */
        private static void ValidateColumn(int column, SpreadsheetVersion ssVersion)
        {
            int maxcol = ssVersion.LastColumnIndex;
            if (column > maxcol) throw new ArgumentException("Maximum column number is " + maxcol);
            if (column < 0) throw new ArgumentException("Minimum column number is 0");
        }
        public bool IsInRange(int rowInd, int colInd)
        {
            return _firstRow <= rowInd && rowInd <= _lastRow &&
                    _firstCol <= colInd && colInd <= _lastCol;
        }
        public bool IsFullColumnRange
        {
            get
            {
                return (_firstRow == 0 && _lastRow ==SpreadsheetVersion.EXCEL97.LastRowIndex)
                    || (_firstRow == -1 && _lastRow == -1);
            }
        }
        public bool IsFullRowRange
        {
            get
            {
                return (_firstCol == 0 && _lastCol == SpreadsheetVersion.EXCEL97.LastColumnIndex)
                    || (_firstCol == -1 && _lastCol == -1);
            }
        }

        /**
         * @return column number for the upper left hand corner
         */
        public int FirstColumn
        {
            get
            {
                return _firstCol;
            }
            set { _firstCol = value; }
        }

        /**
         * @return row number for the upper left hand corner
         */
        public int FirstRow
        {
            get
            {
                return _firstRow;
            }
            set { _firstRow = value; }
        }

        /**
         * @return column number for the lower right hand corner
         */
        public int LastColumn
        {
            get
            {
                return _lastCol;
            }
            set { _lastCol = value; }
        }

        /**
         * @return row number for the lower right hand corner
         */
        public int LastRow
        {
            get
            {
                return _lastRow;
            }
            set { _lastRow = value; }
        }

        /**
         * @return the size of the range (number of cells in the area).
         */
        public int NumberOfCells
        {
            get
            {
                return (_lastRow - _firstRow + 1) * (_lastCol - _firstCol + 1);
            }
        }
        public override String ToString()
        {
            CellReference crA = new CellReference(_firstRow, _firstCol);
            CellReference crB = new CellReference(_lastRow, _lastCol);
            return GetType().Name + " [" + crA.FormatAsString() + ":" + crB.FormatAsString() + "]";
        }
    }
}

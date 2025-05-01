namespace NPOI.SS.Util
{
    using EnumsNET;
    using System;
    using System.Collections.Generic;

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
            if(row > maxrow)
                throw new ArgumentException("Maximum row number is " + maxrow);
            if(row < 0)
                throw new ArgumentException("Minumum row number is 0");
        }

        /**
         * Runs a bounds check for column numbers
         * @param column
         */
        private static void ValidateColumn(int column, SpreadsheetVersion ssVersion)
        {
            int maxcol = ssVersion.LastColumnIndex;
            if(column > maxcol)
                throw new ArgumentException("Maximum column number is " + maxcol);
            if(column < 0)
                throw new ArgumentException("Minimum column number is 0");
        }
        public bool IsInRange(int rowInd, int colInd)
        {
            return _firstRow <= rowInd && rowInd <= _lastRow && //containsRow
                _firstCol <= colInd && colInd <= _lastCol; //containsColumn
        }
        public bool IsInRange(CellReference reference)
        {
            return IsInRange(reference.Row, reference.Col);
        }
        public bool IsFullColumnRange
        {
            get
            {
                return (_firstRow == 0 && _lastRow == SpreadsheetVersion.EXCEL97.LastRowIndex)
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
         * Check if the row is in the specified cell range
         *
         * @param rowInd the row to check
         * @return true if the range contains the row [rowInd]
         */
        public bool ContainsRow(int rowInd)
        {
            return _firstRow <= rowInd && rowInd <= _lastRow;
        }

        /**
         * Check if the column is in the specified cell range
         *
         * @param colInd the column to check
         * @return true if the range contains the column [colInd]
         */
        public bool ContainsColumn(int colInd)
        {
            return _firstCol <= colInd && colInd <= _lastCol;
        }

        /// <summary>
        /// Determines whether or not this CellRangeAddress and the specified CellRangeAddress intersect.
        /// </summary>
        /// <param name="other">a candidate cell range address to check for intersection with this range</param>
        /// <returns>returns true if this range and other range have at least 1 cell in common</returns>
        public bool Intersects(CellRangeAddressBase other)
        {
            return this._firstRow <= other._lastRow &&
                this._firstCol <= other._lastCol &&
                other._firstRow <= this._lastRow &&
                other._firstCol <= this._lastCol;
        }

        /// <summary>
        /// Useful for logic like table/range styling, where some elements apply based on relative position in a range.
        /// </summary>
        /// <param name="rowInd"></param>
        /// <param name="colInd"></param>
        /// <returns>set of <see cref="CellPosition"/>s occupied by the given coordinates.  Empty if the coordinates are not in the range, never null.</returns>
        /// <remarks>
        /// @since 3.17 beta 1
        /// </remarks>
        public ISet<CellPosition> getPosition(int rowInd, int colInd)
        {
	        HashSet<CellPosition> positions = new HashSet<CellPosition>();
	        if (rowInd > FirstRow && rowInd < LastRow && colInd > FirstColumn && colInd < LastColumn) {
	            positions.Add(CellPosition.INSIDE);
	            return positions; // entirely inside, matches no boundaries
	        }
	        // check edges
	        if (rowInd == FirstRow) positions.Add(CellPosition.TOP);
	        if (rowInd == LastRow) positions.Add(CellPosition.BOTTOM);
	        if (colInd == FirstColumn) positions.Add(CellPosition.LEFT);
	        if (colInd == LastColumn) positions.Add(CellPosition.RIGHT);
	    
	        return positions;
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

        // In case _firstRow > _lastRow or _firstCol > _lastCol
        public int MinRow
        {
            get
            {
                return Math.Min(_firstRow, _lastRow);
            }
        }
        public int MaxRow
        {
            get
            {
                return Math.Max(_firstRow, _lastRow);
            }
        }
        public int MinColumn
        {
            get
            {
                return Math.Min(_firstCol, _lastCol);
            }

        }
        public int MaxColumn
        {
            get
            {
                return Math.Max(_firstCol, _lastCol);
            }
        }


        public override bool Equals(Object other)
        {
            if(other is CellRangeAddressBase @base)
            {
                return ((MinRow == @base.MinRow) &&
                        (MaxRow == @base.MaxRow) &&
                        (MinColumn == @base.MinColumn) &&
                        (MaxColumn == @base.MaxColumn));
            }
            return false;
        }

        public override int GetHashCode()
        {
            int code = (MinColumn +
                (MaxColumn << 8) +
                (MinRow << 16) +
                (MaxRow << 24));
            return code;
        }
    }

    /// <summary>
    /// Indicates a cell or range is in the given relative position in a range.
    /// More than one of these may apply at once.
    /// </summary>
    public enum CellPosition
    {
        /// <summary>
        /// range starting rows are equal */
        /// </summary>
        TOP,
        /// <summary>
        /// range ending rows are equal */
        /// </summary>
        BOTTOM,
        /// <summary>
        /// range starting columns are equal */
        /// </summary>
        LEFT,
        /// <summary>
        /// range ending columns are equal */
        /// </summary>
        RIGHT,
        /// <summary>
        /// a cell or range is completely inside another range, without touching any edges (a cell in this position can't be in any others) */
        /// </summary>
        INSIDE
    }
}

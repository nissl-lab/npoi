/*
 * Licensed to the Apache Software Foundation (ASF) Under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for Additional information regarding copyright ownership.
 * The ASF licenses this file to You Under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed Under the License is distributed on an "AS Is" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations Under the License.
 */

namespace NPOI.SS.Formula.Eval
{
    using System;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;


    /**
     * @author Josh Micich
     */
    public abstract class AreaEvalBase : AreaEval
    {
        private int _firstSheet;
        private int _firstColumn;
        private int _firstRow;
        private int _lastSheet;
        private int _lastColumn;
        private int _lastRow;
        private int _nColumns;
        private int _nRows;

        protected AreaEvalBase(ISheetRange sheets, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            _firstColumn = firstColumn;
            _firstRow = firstRow;
            _lastColumn = lastColumn;
            _lastRow = lastRow;

            _nColumns = _lastColumn - _firstColumn + 1;
            _nRows = _lastRow - _firstRow + 1;

            if (sheets != null)
            {
                _firstSheet = sheets.FirstSheetIndex;
                _lastSheet = sheets.LastSheetIndex;
            }
            else
            {
                _firstSheet = -1;
                _lastSheet = -1;
            }
        }
        protected AreaEvalBase(int firstRow, int firstColumn, int lastRow, int lastColumn)
            : this(null, firstRow, firstColumn, lastRow, lastColumn)
        {
        }
        protected AreaEvalBase(AreaI ptg)
            : this(ptg, null)
        {
            
        }
        protected AreaEvalBase(AreaI ptg, ISheetRange sheets)
            : this(sheets, ptg.FirstRow, ptg.FirstColumn, ptg.LastRow, ptg.LastColumn)
        {
            
        }

        public int FirstColumn
        {
            get{return _firstColumn;}
        }

        public int FirstRow
        {
            get { return _firstRow; }
        }

        public int LastColumn
        {
            get { return _lastColumn; }
        }

        public int LastRow
        {
            get { return _lastRow; }
        }

        public int FirstSheetIndex
        {
            get
            {
                return _firstSheet;
            }
        }
        public int LastSheetIndex
        {
            get
            {
                return _lastSheet;
            }
        }

        public ValueEval GetValue(int row, int col)
        {
            return GetRelativeValue(row, col);
        }

        public ValueEval GetValue(int sheetIndex, int row, int col)
        {
            return GetRelativeValue(sheetIndex, row, col);
        }

        public bool Contains(int row, int col)
        {
            return _firstRow <= row && _lastRow >= row
                && _firstColumn <= col && _lastColumn >= col;
        }

        public bool ContainsRow(int row)
        {
            return (_firstRow <= row) && (_lastRow >= row);
        }

        public bool ContainsColumn(int col)
        {
            return (_firstColumn <= col) && (_lastColumn >= col);
        }

        public bool IsColumn
        {
            get{return _firstColumn == _lastColumn;}
        }

        public bool IsRow
        {
            get { return _firstRow == _lastRow; }
        }
        public ValueEval GetAbsoluteValue(int row, int col)
        {
            int rowOffsetIx = row - _firstRow;
            int colOffsetIx = col - _firstColumn;

            if (rowOffsetIx < 0 || rowOffsetIx >= _nRows)
            {
                throw new ArgumentException("Specified row index (" + row
                        + ") is outside the allowed range (" + _firstRow + ".." + _lastRow + ")");
            }
            if (colOffsetIx < 0 || colOffsetIx >= _nColumns)
            {
                throw new ArgumentException("Specified column index (" + col
                        + ") is outside the allowed range (" + _firstColumn + ".." + col + ")");
            }
            return GetRelativeValue(rowOffsetIx, colOffsetIx);
        }
        public abstract ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex);
        public abstract ValueEval GetRelativeValue(int sheetIndex, int relativeRowIndex, int relativeColumnIndex);

        public int Width
        {
            get
            {
                return _lastColumn - _firstColumn + 1;
            }
        }

        public int Height
        {
            get { return _lastRow - _firstRow + 1; }
        }

        /**
 * @return  whether cell at rowIndex and columnIndex is a subtotal.
 * By default return false which means 'don't care about subtotals'
*/
        public virtual bool IsSubTotal(int rowIndex, int columnIndex)
        {
            return false;
        }
        public abstract TwoDEval GetRow(int rowIndex);
        public abstract TwoDEval GetColumn(int columnIndex);

        public abstract AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx);
    }
}
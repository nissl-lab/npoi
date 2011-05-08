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

namespace NPOI.HSSF.Record.Formula.Eval
{
    using System;
    using NPOI.HSSF.Record.Formula;


    /**
     * @author Josh Micich
     */
    public abstract class AreaEvalBase : AreaEval
    {

        private int _firstColumn;
        private int _firstRow;
        private int _lastColumn;
        private int _lastRow;
        private int _nColumns;
        private int _nRows;

        protected AreaEvalBase(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            _firstColumn = firstColumn;
            _firstRow = firstRow;
            _lastColumn = lastColumn;
            _lastRow = lastRow;

            _nColumns = _lastColumn - _firstColumn + 1;
            _nRows = _lastRow - _firstRow + 1;
        }

        protected AreaEvalBase(AreaI ptg)
        {
            _firstRow = ptg.FirstRow;
            _firstColumn = ptg.FirstColumn;
            _lastRow = ptg.LastRow;
            _lastColumn = ptg.LastColumn;

            _nColumns = _lastColumn - _firstColumn + 1;
            _nRows = _lastRow - _firstRow + 1;
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

        public ValueEval GetValueAt(int row, int col)
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

        public abstract ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex);


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

        public abstract AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx);
    }
}
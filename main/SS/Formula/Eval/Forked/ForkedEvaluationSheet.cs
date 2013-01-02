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

using System;
using System.Collections.Generic;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
namespace NPOI.SS.Formula.Eval.Forked
{


    /**
     * Represents a sheet being used for forked Evaluation.  Initially, objects of this class contain
     * only the cells from the master workbook. By calling {@link #getOrCreateUpdatableCell(int, int)},
     * the master cell object is logically Replaced with a {@link ForkedEvaluationCell} instance, which
     * will be used in all subsequent Evaluations.
     *
     * @author Josh Micich
     */
    class ForkedEvaluationSheet : IEvaluationSheet
    {

        private IEvaluationSheet _masterSheet;
        /**
         * Only cells which have been split are Put in this map.  (This has been done to conserve memory).
         */
        private Dictionary<RowColKey, ForkedEvaluationCell> _sharedCellsByRowCol;

        public ForkedEvaluationSheet(IEvaluationSheet masterSheet)
        {
            _masterSheet = masterSheet;
            _sharedCellsByRowCol = new Dictionary<RowColKey, ForkedEvaluationCell>();
        }

        public IEvaluationCell GetCell(int rowIndex, int columnIndex)
        {
            RowColKey key = new RowColKey(rowIndex, columnIndex);

            ForkedEvaluationCell result = null;
            if (_sharedCellsByRowCol.ContainsKey(key))
                result = _sharedCellsByRowCol[(key)];

            if (result == null)
            {
                return _masterSheet.GetCell(rowIndex, columnIndex);
            }
            return result;
        }

        public ForkedEvaluationCell GetOrCreateUpdatableCell(int rowIndex, int columnIndex)
        {
            RowColKey key = new RowColKey(rowIndex, columnIndex);

            ForkedEvaluationCell result = null;
            if (_sharedCellsByRowCol.ContainsKey(key))
                result = _sharedCellsByRowCol[(key)];
            if (result == null)
            {
                IEvaluationCell mcell = _masterSheet.GetCell(rowIndex, columnIndex);
                if (mcell == null)
                {
                    CellReference cr = new CellReference(rowIndex, columnIndex);
                    throw new InvalidOperationException("Underlying cell '"
                            + cr.FormatAsString() + "' is missing in master sheet.");
                }
                result = new ForkedEvaluationCell(this, mcell);
                if (_sharedCellsByRowCol.ContainsKey(key))
                    _sharedCellsByRowCol[key] = result;
                else
                    _sharedCellsByRowCol.Add(key, result);
            }
            return result;
        }

        public void CopyUpdatedCells(ISheet sheet)
        {
            RowColKey[] keys = new RowColKey[_sharedCellsByRowCol.Count];
            _sharedCellsByRowCol.Keys.CopyTo(keys, 0);
            Array.Sort(keys);
            for (int i = 0; i < keys.Length; i++)
            {
                RowColKey key = keys[i];
                IRow row = sheet.GetRow(key.RowIndex);
                if (row == null)
                {
                    row = sheet.CreateRow(key.RowIndex);
                }
                ICell destCell = row.GetCell(key.ColumnIndex);
                if (destCell == null)
                {
                    destCell = row.CreateCell(key.ColumnIndex);
                }

                ForkedEvaluationCell srcCell = _sharedCellsByRowCol[(key)];
                srcCell.CopyValue(destCell);
            }
        }

        public int GetSheetIndex(IEvaluationWorkbook mewb)
        {
            return mewb.GetSheetIndex(_masterSheet);
        }

        private class RowColKey : IComparable<RowColKey>
        {
            private int _rowIndex;
            private int _columnIndex;

            public RowColKey(int rowIndex, int columnIndex)
            {
                _rowIndex = rowIndex;
                _columnIndex = columnIndex;
            }

            public override bool Equals(Object obj)
            {
                //assert obj is RowColKey : "these private cache key instances are only Compared to themselves";
                RowColKey other = (RowColKey)obj;
                return _rowIndex == other._rowIndex && _columnIndex == other._columnIndex;
            }

            public override int GetHashCode()
            {
                return _rowIndex ^ _columnIndex;
            }
            public int CompareTo(RowColKey o)
            {
                int cmp = _rowIndex - o._rowIndex;
                if (cmp != 0)
                {
                    return cmp;
                }
                return _columnIndex - o._columnIndex;
            }
            public int RowIndex
            {
                get
                {
                    return _rowIndex;
                }
            }
            public int ColumnIndex
            {
                get
                {
                    return _columnIndex;
                }
            }
        }
    }

}
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

using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace NPOI.XSSF.UserModel
{

    /**
     * XSSF wrapper for a sheet under Evaluation
     * 
     * @author Josh Micich
     */
    public class XSSFEvaluationSheet : IEvaluationSheet
    {

        private readonly XSSFSheet _xs;
        private Dictionary<CellKey, IEvaluationCell> _cellCache;

        public XSSFEvaluationSheet(ISheet sheet)
        {
            _xs = (XSSFSheet)sheet;
        }

        public XSSFEvaluationSheet()
        {

        }

        public void ClearAllCachedResultValues()
        {
            _cellCache = null;
        }

        public XSSFSheet XSSFSheet => _xs;
        
        public IEvaluationCell GetCell(int rowIndex, int columnIndex)
        {
            // cache for performance: ~30% speedup due to caching
            if (_cellCache == null)
            {
                _cellCache = new Dictionary<CellKey, IEvaluationCell>(_xs.LastRowNum * 3);
                foreach (IRow row in _xs)
                {
                    int rowNum = row.RowNum;
                    foreach (ICell cell in row)
                    {
                        // cast is safe, the iterator is just defined using the interface
                        CellKey key1 = new CellKey(rowNum, cell.ColumnIndex);
                        IEvaluationCell evalcell1 = new XSSFEvaluationCell((XSSFCell)cell, this);
                        _cellCache.Add(key1, evalcell1);
                    }
                }
            }

            CellKey key = new CellKey(rowIndex, columnIndex);

            IEvaluationCell evalcell = null;
            if (_cellCache.TryGetValue(key, out IEvaluationCell value))
            {
                evalcell = value;           
            }           

            // If cache is stale, update cache with this one cell
            // This is a compromise between rebuilding the entire cache
            // (which would quickly defeat the benefit of the cache)
            // and not caching at all.
            // See bug 59958: Add cells on the fly to the evaluation sheet cache on cache miss
            if (evalcell == null)
            {
                XSSFRow row = _xs.GetRow(rowIndex) as XSSFRow;
                if (row == null)
                {
                    return null;
                }
                XSSFCell cell = row.GetCell(columnIndex) as XSSFCell;
                if (cell == null)
                {
                    return null;
                }
                evalcell = new XSSFEvaluationCell(cell, this);
                _cellCache[key] = evalcell;
            }

            return evalcell;
        }

        private class CellKey
        {
            private readonly int _row;
            private readonly int _col;
            private int _hash = -1; //lazily computed

            protected internal CellKey(int row, int col)
            {
                _row = row;
                _col = col;
            }

            public override int GetHashCode()
            {
                if (_hash == -1)
                {
                    _hash = (17 * 37 + _row) * 37 + _col;
                }
                return _hash;
            }

            public override bool Equals(Object obj)
            {
                if (obj == null || !(obj is CellKey))
                    return false;

                // assumes other object is one of us, otherwise ClassCastException is thrown
                CellKey oKey = (CellKey)obj;
                return _row == oKey._row && _col == oKey._col;
            }
        }
    }
}


/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using System.Collections;
    using System.Text;
    using NPOI.SS.Formula.Eval;

    /**
     * Stores details about the current evaluation of a cell.<br/>
     */
    class CellEvaluationFrame
    {

        private FormulaCellCacheEntry _cce;
        private ArrayList _sensitiveInputCells;
        private FormulaUsedBlankCellSet _usedBlankCellGroup;

        public CellEvaluationFrame(FormulaCellCacheEntry cce)
        {
            _cce = cce;
            _sensitiveInputCells = new ArrayList();
        }
        public CellCacheEntry GetCCE()
        {
            return _cce;
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append("]");
            return sb.ToString();
        }
        /**
         * @param inputCell a cell directly used by the formula of this evaluation frame
         */
        public void AddSensitiveInputCell(CellCacheEntry inputCell)
        {
            _sensitiveInputCells.Add(inputCell);
        }
        /**
         * @return never <c>null</c>, (possibly empty) array of all cells directly used while 
         * evaluating the formula of this frame.
         */
        private CellCacheEntry[] GetSensitiveInputCells()
        {
            int nItems = _sensitiveInputCells.Count;
            if (nItems < 1)
            {
                return CellCacheEntry.EMPTY_ARRAY;
            }
            CellCacheEntry[] result = new CellCacheEntry[nItems];
            result = (CellCacheEntry[])_sensitiveInputCells.ToArray(typeof(CellCacheEntry));
            return result;
        }
        public void AddUsedBlankCell(int bookIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            if (_usedBlankCellGroup == null)
            {
                _usedBlankCellGroup = new FormulaUsedBlankCellSet();
            }
            _usedBlankCellGroup.AddCell(bookIndex, sheetIndex, rowIndex, columnIndex);
        }

        public void UpdateFormulaResult(ValueEval result)
        {
            _cce.UpdateFormulaResult(result, GetSensitiveInputCells(), _usedBlankCellGroup);
        }
    }
}
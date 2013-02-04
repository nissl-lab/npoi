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
    using System.Collections;
    using NPOI.SS.Formula.Eval;


    /**
     * Stores the cached result of a formula evaluation, along with the Set of sensititive input cells
     * 
     * @author Josh Micich
     */
    public class FormulaCellCacheEntry : CellCacheEntry
    {
        public static new FormulaCellCacheEntry[] EMPTY_ARRAY = { };

        /**
         * Cells 'used' in the current evaluation of the formula corresponding To this cache entry
         *
         * If any of the following cells Change, this cache entry needs To be Cleared
         */
        private CellCacheEntry[] _sensitiveInputCells;

        private FormulaUsedBlankCellSet _usedBlankCellGroup;

        public FormulaCellCacheEntry()
        {

        }

        public bool IsInputSensitive
        {
            get
            {
                if (_sensitiveInputCells != null)
                {
                    if (_sensitiveInputCells.Length > 0)
                    {
                        return true;
                    }
                }
                return _usedBlankCellGroup == null ? false : !_usedBlankCellGroup.IsEmpty;
            }
        }

        public void SetSensitiveInputCells(CellCacheEntry[] sensitiveInputCells)
        {
            // need To tell all cells that were previously used, but no longer are, 
            // that they are not consumed by this cell any more
            ChangeConsumingCells(sensitiveInputCells == null ? CellCacheEntry.EMPTY_ARRAY : sensitiveInputCells);
            _sensitiveInputCells = sensitiveInputCells;
        }

        public void ClearFormulaEntry()
        {
            CellCacheEntry[] usedCells = _sensitiveInputCells;
            if (usedCells != null)
            {
                for (int i = usedCells.Length - 1; i >= 0; i--)
                {
                    usedCells[i].ClearConsumingCell(this);
                }
            }
            _sensitiveInputCells = null;
            ClearValue();
        }

        private void ChangeConsumingCells(CellCacheEntry[] usedCells)
        {

            CellCacheEntry[] prevUsedCells = _sensitiveInputCells;
            int nUsed = usedCells.Length;
            for (int i = 0; i < nUsed; i++)
            {
                usedCells[i].AddConsumingCell(this);
            }
            if (prevUsedCells == null)
            {
                return;
            }
            int nPrevUsed = prevUsedCells.Length;
            if (nPrevUsed < 1)
            {
                return;
            }
            ArrayList usedSet;
            if (nUsed < 1)
            {
                usedSet = new ArrayList();
            }
            else
            {
                usedSet = new ArrayList(nUsed * 3 / 2);
                for (int i = 0; i < nUsed; i++)
                {
                    usedSet.Add(usedCells[i]);
                }
            }
            for (int i = 0; i < nPrevUsed; i++)
            {
                CellCacheEntry prevUsed = prevUsedCells[i];
                if (!usedSet.Contains(prevUsed))
                {
                    // previously was used by cellLoc, but not anymore
                    prevUsed.ClearConsumingCell(this);
                }
            }
        }

        public void UpdateFormulaResult(ValueEval result, CellCacheEntry[] sensitiveInputCells, FormulaUsedBlankCellSet usedBlankAreas)
        {
            UpdateValue(result);
            SetSensitiveInputCells(sensitiveInputCells);
            _usedBlankCellGroup = usedBlankAreas;
        }

        public void NotifyUpdatedBlankCell(BookSheetKey bsk, int rowIndex, int columnIndex, IEvaluationListener evaluationListener)
        {
            if (_usedBlankCellGroup != null)
            {
                if (_usedBlankCellGroup.ContainsCell(bsk, rowIndex, columnIndex))
                {
                    ClearFormulaEntry();
                    RecurseClearCachedFormulaResults(evaluationListener);
                }
            }
        }
    }
}
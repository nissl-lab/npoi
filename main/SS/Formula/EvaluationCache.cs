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

    using NPOI.SS.Formula.Eval;
    

    /**
     * Performance optimisation for {@link HSSFFormulaEvaluator}. This class stores previously
     * calculated values of already visited cells, To avoid unnecessary re-calculation when the 
     * same cells are referenced multiple times
     * 
     * 
     * @author Josh Micich
     */
    public class EvaluationCache
    {

        private PlainCellCache _plainCellCache;
        private FormulaCellCache _formulaCellCache;
        /** only used for testing. <c>null</c> otherwise */
        IEvaluationListener _evaluationListener;

        /* package */
        public EvaluationCache(IEvaluationListener evaluationListener)
        {
            _evaluationListener = evaluationListener;
            _plainCellCache = new PlainCellCache();
            _formulaCellCache = new FormulaCellCache();
        }

        public void NotifyUpdateCell(int bookIndex, int sheetIndex, IEvaluationCell cell)
        {
            FormulaCellCacheEntry fcce = _formulaCellCache.Get(cell);

            int rowIndex = cell.RowIndex;
            int columnIndex = cell.ColumnIndex;
            Loc loc = new Loc(bookIndex, sheetIndex, cell.RowIndex, cell.ColumnIndex);
            PlainValueCellCacheEntry pcce = _plainCellCache.Get(loc);

            if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
            {
                if (fcce == null)
                {
                    fcce = new FormulaCellCacheEntry();

                    if (pcce == null)
                    {
                        if (_evaluationListener != null)
                        {
                            _evaluationListener.OnChangeFromBlankValue(sheetIndex, rowIndex,
                                    columnIndex, cell, fcce);
                        }
                        UpdateAnyBlankReferencingFormulas(bookIndex, sheetIndex, rowIndex, columnIndex);
                    }
                    _formulaCellCache.Put(cell, fcce);
                }
                else
                {
                    fcce.RecurseClearCachedFormulaResults(_evaluationListener);
                    fcce.ClearFormulaEntry();
                }
                if (pcce == null)
                {
                    // was formula cell before - no Change of type
                }
                else
                {
                    // changing from plain cell To formula cell
                    pcce.RecurseClearCachedFormulaResults(_evaluationListener);
                    _plainCellCache.Remove(loc);
                }
            }
            else
            {
                ValueEval value = WorkbookEvaluator.GetValueFromNonFormulaCell(cell);
                if (pcce == null)
                {
                    if (value != BlankEval.instance)
                    {
                        pcce = new PlainValueCellCacheEntry(value);
                        if (fcce == null)
                        {
                            if (_evaluationListener != null)
                            {
                                _evaluationListener.OnChangeFromBlankValue(sheetIndex, rowIndex,
                                        columnIndex, cell, pcce);
                            }
                            UpdateAnyBlankReferencingFormulas(bookIndex, sheetIndex, rowIndex, columnIndex);
                        }
                        _plainCellCache.Put(loc, pcce);
                    }
                }
                else
                {
                    if (pcce.UpdateValue(value))
                    {
                        pcce.RecurseClearCachedFormulaResults(_evaluationListener);
                    }
                    if (value == BlankEval.instance)
                    {
                        _plainCellCache.Remove(loc);
                    }
                }
                if (fcce == null)
                {
                    // was plain cell before - no Change of type
                }
                else
                {
                    // was formula cell before - now a plain value
                    _formulaCellCache.Remove(cell);
                    fcce.SetSensitiveInputCells(null);
                    fcce.RecurseClearCachedFormulaResults(_evaluationListener);
                }
            }
        }

        public class EntryOperation : IEntryOperation
        {
            BookSheetKey bsk;
            int rowIndex, columnIndex;
            IEvaluationListener evaluationListener;

            public EntryOperation(BookSheetKey bsk,
                int rowIndex, int columnIndex, IEvaluationListener evaluationListener)
            {
                this.bsk = bsk;
                this.evaluationListener = evaluationListener;
                this.rowIndex = rowIndex;
                this.columnIndex = columnIndex;
            }

            public void ProcessEntry(FormulaCellCacheEntry entry)
            {
                entry.NotifyUpdatedBlankCell(bsk, rowIndex, columnIndex, evaluationListener);
            }
        }

        private void UpdateAnyBlankReferencingFormulas(int bookIndex, int sheetIndex,
                int rowIndex, int columnIndex)
        {
            BookSheetKey bsk = new BookSheetKey(bookIndex, sheetIndex);
            _formulaCellCache.ApplyOperation(new EntryOperation(bsk,rowIndex,columnIndex,_evaluationListener));
        }

        public PlainValueCellCacheEntry GetPlainValueEntry(int bookIndex, int sheetIndex,
                int rowIndex, int columnIndex, ValueEval value)
        {

            Loc loc = new Loc(bookIndex, sheetIndex, rowIndex, columnIndex);
            PlainValueCellCacheEntry result = _plainCellCache.Get(loc);
            if (result == null)
            {
                result = new PlainValueCellCacheEntry(value);
                _plainCellCache.Put(loc, result);
                if (_evaluationListener != null)
                {
                    _evaluationListener.OnReadPlainValue(sheetIndex, rowIndex, columnIndex, result);
                }
            }
            else
            {
                // TODO - if we are confident that this sanity check is not required, we can Remove 'value' from plain value cache entry  
                if (!AreValuesEqual(result.GetValue(), value))
                {
                    throw new InvalidOperationException("value changed");
                }
                if (_evaluationListener != null)
                {
                    _evaluationListener.OnCacheHit(sheetIndex, rowIndex, columnIndex, value);
                }
            }
            return result;
        }
        private bool AreValuesEqual(ValueEval a, ValueEval b)
        {
            if (a == null)
            {
                return false;
            }
            Type cls = a.GetType();
            if (cls != b.GetType())
            {
                // value type is changing
                return false;
            }
            if (a == BlankEval.instance)
            {
                return b == a;
            }
            if (cls == typeof(NumberEval))
            {
                return ((NumberEval)a).NumberValue == ((NumberEval)b).NumberValue;
            }
            if (cls == typeof(StringEval))
            {
                return ((StringEval)a).StringValue.Equals(((StringEval)b).StringValue);
            }
            if (cls == typeof(BoolEval))
            {
                return ((BoolEval)a).BooleanValue == ((BoolEval)b).BooleanValue;
            }
            if (cls == typeof(ErrorEval))
            {
                return ((ErrorEval)a).ErrorCode == ((ErrorEval)b).ErrorCode;
            }
            throw new InvalidOperationException("Unexpected value class (" + cls.Name + ")");
        }

        public FormulaCellCacheEntry GetOrCreateFormulaCellEntry(IEvaluationCell cell)
        {            
            FormulaCellCacheEntry result = _formulaCellCache.Get(cell);
            if (result == null)
            {

                result = new FormulaCellCacheEntry();
                _formulaCellCache.Put(cell, result);
            }
            return result;
        }

        /**
         * Should be called whenever there are Changes To input cells in the evaluated workbook.
         */
        public void Clear()
        {
            if (_evaluationListener != null)
            {
                _evaluationListener.OnClearWholeCache();
            }
            _plainCellCache.Clear();
            _formulaCellCache.Clear();
        }
        public void NotifyDeleteCell(int bookIndex, int sheetIndex, IEvaluationCell cell)
        {

            if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
            {
                FormulaCellCacheEntry fcce = _formulaCellCache.Remove(cell);
                if (fcce == null)
                {
                    // formula cell Has not been evaluated yet 
                }
                else
                {
                    fcce.SetSensitiveInputCells(null);
                    fcce.RecurseClearCachedFormulaResults(_evaluationListener);
                }
            }
            else
            {
                Loc loc = new Loc(bookIndex, sheetIndex, cell.RowIndex, cell.ColumnIndex);
                PlainValueCellCacheEntry pcce = _plainCellCache.Get(loc);

                if (pcce == null)
                {
                    // cache entry doesn't exist. nothing To do
                }
                else
                {
                    pcce.RecurseClearCachedFormulaResults(_evaluationListener);
                }
            }
        }
    }
}
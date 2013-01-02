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
     * Stores the parameters that identify the evaluation of one cell.<br/>
     */
    public abstract class CellCacheEntry : ICacheEntry
    {
        public static CellCacheEntry[] EMPTY_ARRAY = { };

        private FormulaCellCacheEntrySet _consumingCells;
        private ValueEval _value;


        protected CellCacheEntry()
        {
            _consumingCells = new FormulaCellCacheEntrySet();
        }
        protected void ClearValue()
        {
            _value = null;
        }

        public bool UpdateValue(ValueEval value)
        {
            if (value == null)
            {
                throw new ArgumentException("Did not expect To Update To null");
            }
            bool result = !AreValuesEqual(_value, value);
            _value = value;
            return result;
        }
        public ValueEval GetValue()
        {
            return _value;
        }

        private static bool AreValuesEqual(ValueEval a, ValueEval b)
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

        public void AddConsumingCell(FormulaCellCacheEntry cellLoc)
        {
            _consumingCells.Add(cellLoc);

        }
        public FormulaCellCacheEntry[] GetConsumingCells()
        {
            return _consumingCells.ToArray();
        }

        public void ClearConsumingCell(FormulaCellCacheEntry cce)
        {
            if (!_consumingCells.Remove(cce))
            {
                throw new InvalidOperationException("Specified formula cell is not consumed by this cell");
            }
        }
        public void RecurseClearCachedFormulaResults(IEvaluationListener listener)
        {
            if (listener == null)
            {
                RecurseClearCachedFormulaResults();
            }
            else
            {
                listener.OnClearCachedValue(this);
                RecurseClearCachedFormulaResults(listener, 1);
            }
        }

        /**
         * Calls formulaCell.SetFormulaResult(null, null) recursively all the way up the tree of 
         * dependencies. Calls usedCell.ClearConsumingCell(fc) for each child of a cell that Is
         * Cleared along the way.
         * @param formulaCells
         */
        protected void RecurseClearCachedFormulaResults()
        {
            FormulaCellCacheEntry[] formulaCells = GetConsumingCells();

            for (int i = 0; i < formulaCells.Length; i++)
            {
                FormulaCellCacheEntry fc = formulaCells[i];
                fc.ClearFormulaEntry();
                fc.RecurseClearCachedFormulaResults();
            }
        }

        /**
         * Identical To {@link #RecurseClearCachedFormulaResults()} except for the listener call-backs
         */
        protected void RecurseClearCachedFormulaResults(IEvaluationListener listener, int depth)
        {
            FormulaCellCacheEntry[] formulaCells = GetConsumingCells();

            listener.SortDependentCachedValues(formulaCells);
            for (int i = 0; i < formulaCells.Length; i++)
            {
                FormulaCellCacheEntry fc = formulaCells[i];
                listener.OnClearDependentCachedValue(fc, depth);
                fc.ClearFormulaEntry();
                fc.RecurseClearCachedFormulaResults(listener, depth + 1);
            }
        }

    }
}
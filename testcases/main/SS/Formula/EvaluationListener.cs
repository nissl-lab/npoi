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

namespace TestCases.SS.Formula
{
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;

    /**
     * Tests should extend this class if they need to track the internal working of the {@link WorkbookEvaluator}.<br/>
     *
     * Default method implementations all do nothing
     *
     * @author Josh Micich
     */
    public abstract class EvaluationListener : IEvaluationListener
    {
        public virtual void OnCacheHit(int sheetIndex, int rowIndex, int columnIndex, ValueEval result)
        {
            // do nothing
        }
        public virtual void OnReadPlainValue(int sheetIndex, int rowIndex, int columnIndex, ICacheEntry entry)
        {
            // do nothing
        }
        public virtual void OnStartEvaluate(IEvaluationCell cell, ICacheEntry entry)
        {
            // do nothing
        }
        public virtual void OnEndEvaluate(ICacheEntry entry, ValueEval result)
        {
            // do nothing
        }
        public virtual void OnClearWholeCache()
        {
            // do nothing
        }
        public virtual void OnClearCachedValue(ICacheEntry entry)
        {
            // do nothing
        }
        public virtual void OnChangeFromBlankValue(int sheetIndex, int rowIndex, int columnIndex,
                IEvaluationCell cell, ICacheEntry entry)
        {
            // do nothing
        }
        public virtual void SortDependentCachedValues(ICacheEntry[] entries)
        {
            // do nothing
        }
        public virtual void OnClearDependentCachedValue(ICacheEntry entry, int depth)
        {
            // do nothing
        }
    }

}
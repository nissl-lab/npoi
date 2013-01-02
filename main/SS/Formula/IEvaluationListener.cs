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

    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    

    /**
 * A (mostly) opaque interface To allow test clients To trace cache values
 * Each spreadsheet cell Gets one unique cache entry instance.  These objects
 * are safe To use as keys in {@link java.util.HashMap}s 
 */
    public interface ICacheEntry
    {
        ValueEval GetValue();
    }

    /**
     * Tests can implement this class To track the internal working of the {@link WorkbookEvaluator}.<br/>
     * 
     * For POI internal testing use only
     * 
     * @author Josh Micich
     */
    public interface IEvaluationListener
    {


        void OnCacheHit(int sheetIndex, int rowIndex, int columnIndex, ValueEval result);
        void OnReadPlainValue(int sheetIndex, int rowIndex, int columnIndex, ICacheEntry entry);
        void OnStartEvaluate(IEvaluationCell cell, ICacheEntry entry);
        void OnEndEvaluate(ICacheEntry entry, ValueEval result);
        void OnClearWholeCache();
        void OnClearCachedValue(ICacheEntry entry);
        /**
         * Internally, formula {@link ICacheEntry}s are stored in Sets which may Change ordering due 
         * To seemingly trivial Changes.  This method is provided To make the order of call-backs To 
         * {@link #onClearDependentCachedValue(ICacheEntry, int)} more deterministic.
         */
        void SortDependentCachedValues(ICacheEntry[] formulaCells);
        void OnClearDependentCachedValue(ICacheEntry formulaCell, int depth);
        void OnChangeFromBlankValue(int sheetIndex, int rowIndex, int columnIndex,IEvaluationCell cell, ICacheEntry entry);
    }
}
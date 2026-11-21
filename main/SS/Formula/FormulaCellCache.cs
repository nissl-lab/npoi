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
    using System.Collections.Generic;

    public interface IEntryOperation
    {
        void ProcessEntry(FormulaCellCacheEntry entry);
    }

    /**
     * 
     * @author Josh Micich
     */
    public class FormulaCellCache
    {

        private readonly Dictionary<object, FormulaCellCacheEntry> _formulaEntriesByCell;

        public FormulaCellCache()
        {
            // assumes HSSFCell does not override HashCode or Equals, otherwise we need IdentityHashMap
            _formulaEntriesByCell = new Dictionary<object, FormulaCellCacheEntry>();
        }

        public CellCacheEntry[] GetCacheEntries()
        {
            FormulaCellCacheEntry[] result = new FormulaCellCacheEntry[_formulaEntriesByCell.Count];
            _formulaEntriesByCell.Values.CopyTo(result, 0);
            return result;
        }

        public void Clear()
        {
            _formulaEntriesByCell.Clear();
        }

        /**
         * @return <c>null</c> if not found
         */
        public FormulaCellCacheEntry Get(IEvaluationCell cell)
        {
            _formulaEntriesByCell.TryGetValue(cell.IdentityKey, out FormulaCellCacheEntry entry);
            return entry;
        }

        public void Put(IEvaluationCell cell, FormulaCellCacheEntry entry)
        {
            _formulaEntriesByCell[cell.IdentityKey] = entry;
        }

        public FormulaCellCacheEntry Remove(IEvaluationCell cell)
        {
            _formulaEntriesByCell.TryGetValue(cell.IdentityKey, out FormulaCellCacheEntry tmp);
            _formulaEntriesByCell.Remove(cell.IdentityKey);
            return tmp;
        }

        public void ApplyOperation(IEntryOperation operation)
        {
            foreach (FormulaCellCacheEntry entry in _formulaEntriesByCell.Values)
            {
                operation.ProcessEntry(entry);
            }
        }
    }
}
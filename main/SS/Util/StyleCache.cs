/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.Util
{
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Per-workbook O(1) style lookup table backed by a
    /// <see cref="ConditionalWeakTable{TKey,TValue}"/> so each workbook
    /// automatically gets its own cache that is GC'd with the workbook.
    /// </summary>
    internal sealed class StyleCache
    {
        private static readonly ConditionalWeakTable<IWorkbook, StyleCache> s_table =
            new ConditionalWeakTable<IWorkbook, StyleCache>();

        private readonly Dictionary<StyleKey, ICellStyle> _map =
            new Dictionary<StyleKey, ICellStyle>();

        /// <summary>
        /// Returns the <see cref="StyleCache"/> for the given workbook,
        /// creating it on first access.
        /// </summary>
        public static StyleCache ForWorkbook(IWorkbook wb)
        {
            return s_table.GetValue(wb, _ => new StyleCache());
        }

        /// <summary>
        /// Looks up a style by key.
        /// </summary>
        public bool TryGet(in StyleKey key, out ICellStyle style)
        {
            return _map.TryGetValue(key, out style);
        }

        /// <summary>
        /// Registers a newly created style in the cache.
        /// </summary>
        public void Register(in StyleKey key, ICellStyle style)
        {
            _map[key] = style;
        }

        /// <summary>
        /// Seeds the cache from all styles already present in the workbook.
        /// The first style encountered for a given key wins (no overwrite).
        /// Call this optionally after opening an existing workbook to allow
        /// the cache to reuse pre-existing styles immediately.
        /// </summary>
        public void Warm(IWorkbook wb)
        {
            int count = wb.NumCellStyles;
            for (int i = 0; i < count; i++)
            {
                ICellStyle style = wb.GetCellStyleAt(i);
                StyleKey key = StyleKey.From(style);
                if (!_map.ContainsKey(key))
                    _map[key] = style;
            }
        }

        /// <summary>
        /// Removes a single entry from the cache (e.g. after an
        /// <c>HSSFOptimiser</c> pass that may have changed style indices).
        /// </summary>
        public void Invalidate(in StyleKey key)
        {
            _map.Remove(key);
        }

        /// <summary>
        /// Clears all entries from the cache.
        /// </summary>
        public void Clear()
        {
            _map.Clear();
        }
    }
}

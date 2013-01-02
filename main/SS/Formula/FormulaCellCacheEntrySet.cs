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

    /**
     * A custom implementation of {@link java.util.HashSet} in order To reduce memory consumption.
     *
     * Profiling tests (Oct 2008) have shown that each element {@link FormulaCellCacheEntry} takes
     * around 32 bytes To store in a HashSet, but around 6 bytes To store here.  For Spreadsheets with
     * thousands of formula cells with multiple interdependencies, the savings can be very significant.
     *
     * @author Josh Micich
     */
    class FormulaCellCacheEntrySet
    {

        private int _size;
        private FormulaCellCacheEntry[] _arr;

        public FormulaCellCacheEntrySet()
        {
            _arr = FormulaCellCacheEntry.EMPTY_ARRAY;
        }

        public FormulaCellCacheEntry[] ToArray()
        {
            int nItems = _size;
            if (nItems < 1)
            {
                return FormulaCellCacheEntry.EMPTY_ARRAY;
            }
            FormulaCellCacheEntry[] result = new FormulaCellCacheEntry[nItems];
            int j = 0;
            for (int i = 0; i < _arr.Length; i++)
            {
                FormulaCellCacheEntry cce = _arr[i];
                if (cce != null)
                {
                    result[j++] = cce;
                }
            }
            if (j != nItems)
            {
                throw new InvalidOperationException("size mismatch");
            }
            return result;
        }


        public void Add(CellCacheEntry cce)
        {
            if (_size * 3 >= _arr.Length * 2)
            {
                // re-Hash
                FormulaCellCacheEntry[] prevArr = _arr;
                FormulaCellCacheEntry[] newArr = new FormulaCellCacheEntry[4 + _arr.Length * 3 / 2]; // grow 50%
                for (int i = 0; i < prevArr.Length; i++)
                {
                    FormulaCellCacheEntry prevCce = _arr[i];
                    if (prevCce != null)
                    {
                        AddInternal(newArr, prevCce);
                    }
                }
                _arr = newArr;
            }
            if (AddInternal(_arr, cce))
            {
                _size++;
            }
        }


        private static bool AddInternal(CellCacheEntry[] arr, CellCacheEntry cce)
        {

            int startIx = cce.GetHashCode() % arr.Length;

            for (int i = startIx; i < arr.Length; i++)
            {
                CellCacheEntry item = arr[i];
                if (item == cce)
                {
                    // already present
                    return false;
                }
                if (item == null)
                {
                    arr[i] = cce;
                    return true;
                }
            }
            for (int i = 0; i < startIx; i++)
            {
                CellCacheEntry item = arr[i];
                if (item == cce)
                {
                    // already present
                    return false;
                }
                if (item == null)
                {
                    arr[i] = cce;
                    return true;
                }
            }
            throw new InvalidOperationException("No empty space found");
        }

        public bool Remove(CellCacheEntry cce)
        {
            FormulaCellCacheEntry[] arr = _arr;

            if (_size * 3 < _arr.Length && _arr.Length > 8)
            {
                // re-Hash
                bool found = false;
                FormulaCellCacheEntry[] prevArr = _arr;
                FormulaCellCacheEntry[] newArr = new FormulaCellCacheEntry[_arr.Length / 2]; // shrink 50%
                for (int i = 0; i < prevArr.Length; i++)
                {
                    FormulaCellCacheEntry prevCce = _arr[i];
                    if (prevCce != null)
                    {
                        if (prevCce == cce)
                        {
                            found = true;
                            _size--;
                            // skip it
                            continue;
                        }
                        AddInternal(newArr, prevCce);
                    }
                }
                _arr = newArr;
                return found;
            }
            // else - usual case
            // delete single element (without re-Hashing)

            int startIx = cce.GetHashCode() % arr.Length;

            // note - can't exit loops upon finding null because of potential previous deletes
            for (int i = startIx; i < arr.Length; i++)
            {
                FormulaCellCacheEntry item = arr[i];
                if (item == cce)
                {
                    // found it
                    arr[i] = null;
                    _size--;
                    return true;
                }
            }
            for (int i = 0; i < startIx; i++)
            {
                FormulaCellCacheEntry item = arr[i];
                if (item == cce)
                {
                    // found it
                    arr[i] = null;
                    _size--;
                    return true;
                }
            }
            return false;
        }

    }
}
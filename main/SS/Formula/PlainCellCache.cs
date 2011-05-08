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

    public class Loc
    {

        private int _bookSheetColumn;

        private int _rowIndex;

        public Loc(int bookIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            _bookSheetColumn = ToBookSheetColumn(bookIndex, sheetIndex, columnIndex);
            _rowIndex = rowIndex;
        }

        public static int ToBookSheetColumn(int bookIndex, int sheetIndex, int columnIndex)
        {
            return ((bookIndex & 0x00FF) << 24) + ((sheetIndex & 0x00FF) << 16)
                    + ((columnIndex & 0xFFFF) << 0);
        }

        public Loc(int bookSheetColumn, int rowIndex)
        {
            _bookSheetColumn = bookSheetColumn;
            _rowIndex = rowIndex;
        }

        public override int GetHashCode()
        {
            return _bookSheetColumn + 17 * _rowIndex;
        }

        public override bool Equals(Object obj)
        {
            Loc other = (Loc)obj;
            return _bookSheetColumn == other._bookSheetColumn && _rowIndex == other._rowIndex;
        }

        public int RowIndex
        {
            get
            {
                return _rowIndex;
            }
        }
        public int ColumnIndex
        {
            get
            {
                return _bookSheetColumn & 0x000FFFF;
            }
        }
    }
    /**
     *
     * @author Josh Micich
     */
    class PlainCellCache
    {

        private Hashtable _plainValueEntriesByLoc;

        public PlainCellCache()
        {
            _plainValueEntriesByLoc = new Hashtable();
        }
        public void Put(Loc key, PlainValueCellCacheEntry cce)
        {
            _plainValueEntriesByLoc[key] = cce;
        }
        public void Clear()
        {
            _plainValueEntriesByLoc.Clear();
        }
        public PlainValueCellCacheEntry Get(Loc key)
        {
            return (PlainValueCellCacheEntry)_plainValueEntriesByLoc[key];
        }
        public void Remove(Loc key)
        {
            _plainValueEntriesByLoc.Remove(key);
        }
    }
}
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
    using System.Text;
    using System.Collections;
    using NPOI.SS.Util;

    public class BookSheetKey
    {

        private int _bookIndex;
        private int _sheetIndex;

        public BookSheetKey(int bookIndex, int sheetIndex)
        {
            _bookIndex = bookIndex;
            _sheetIndex = sheetIndex;
        }
        public override int GetHashCode()
        {
            return _bookIndex * 17 + _sheetIndex;
        }
        public override bool Equals(Object obj)
        {
            BookSheetKey other = (BookSheetKey)obj;
            return _bookIndex == other._bookIndex && _sheetIndex == other._sheetIndex;
        }
    }
    /**
     * Optimisation - compacts many blank cell references used by a single formula. 
     * 
     * @author Josh Micich
     */
    public class FormulaUsedBlankCellSet
    {


        private  class BlankCellSheetGroup
        {
            private IList _rectangleGroups;
            private int _currentRowIndex;
            private int _firstColumnIndex;
            private int _lastColumnIndex;
            private BlankCellRectangleGroup _currentRectangleGroup;

            public BlankCellSheetGroup()
            {
                _rectangleGroups = new ArrayList();
                _currentRowIndex = -1;
            }

            public void AddCell(int rowIndex, int columnIndex)
            {
                if (_currentRowIndex == -1)
                {
                    _currentRowIndex = rowIndex;
                    _firstColumnIndex = columnIndex;
                    _lastColumnIndex = columnIndex;
                }
                else
                {
                    if (_currentRowIndex == rowIndex && _lastColumnIndex + 1 == columnIndex)
                    {
                        _lastColumnIndex = columnIndex;
                    }
                    else
                    {
                        // cell does not fit on end of current row
                        if (_currentRectangleGroup == null)
                        {
                            _currentRectangleGroup = new BlankCellRectangleGroup(_currentRowIndex, _firstColumnIndex, _lastColumnIndex);
                        }
                        else
                        {
                            if (!_currentRectangleGroup.AcceptRow(_currentRowIndex, _firstColumnIndex, _lastColumnIndex))
                            {
                                _rectangleGroups.Add(_currentRectangleGroup);
                                _currentRectangleGroup = new BlankCellRectangleGroup(_currentRowIndex, _firstColumnIndex, _lastColumnIndex);
                            }
                        }
                        _currentRowIndex = rowIndex;
                        _firstColumnIndex = columnIndex;
                        _lastColumnIndex = columnIndex;
                    }
                }
            }

            public bool ContainsCell(int rowIndex, int columnIndex)
            {
                for (int i = _rectangleGroups.Count - 1; i >= 0; i--)
                {
                    BlankCellRectangleGroup bcrg = (BlankCellRectangleGroup)_rectangleGroups[i];
                    if (bcrg.ContainsCell(rowIndex, columnIndex))
                    {
                        return true;
                    }
                }
                if (_currentRectangleGroup != null && _currentRectangleGroup.ContainsCell(rowIndex, columnIndex))
                {
                    return true;
                }
                if (_currentRowIndex != -1 && _currentRowIndex == rowIndex)
                {
                    if (_firstColumnIndex <= columnIndex && columnIndex <= _lastColumnIndex)
                    {
                        return true;
                    }
                }
                return false;
            }

        }

        private class BlankCellRectangleGroup
        {

            private int _firstRowIndex;
            private int _firstColumnIndex;
            private int _lastColumnIndex;
            private int _lastRowIndex;

            public BlankCellRectangleGroup(int firstRowIndex, int firstColumnIndex, int lastColumnIndex)
            {
                _firstRowIndex = firstRowIndex;
                _firstColumnIndex = firstColumnIndex;
                _lastColumnIndex = lastColumnIndex;
                _lastRowIndex = firstRowIndex;
            }

            public bool ContainsCell(int rowIndex, int columnIndex)
            {
                if (columnIndex < _firstColumnIndex)
                {
                    return false;
                }
                if (columnIndex > _lastColumnIndex)
                {
                    return false;
                }
                if (rowIndex < _firstRowIndex)
                {
                    return false;
                }
                if (rowIndex > _lastRowIndex)
                {
                    return false;
                }
                return true;
            }

            public bool AcceptRow(int rowIndex, int firstColumnIndex, int lastColumnIndex)
            {
                if (firstColumnIndex != _firstColumnIndex)
                {
                    return false;
                }
                if (lastColumnIndex != _lastColumnIndex)
                {
                    return false;
                }
                if (rowIndex != _lastRowIndex + 1)
                {
                    return false;
                }
                _lastRowIndex = rowIndex;
                return true;
            }
            public override String ToString()
            {
                StringBuilder sb = new StringBuilder(64);
                CellReference crA = new CellReference(_firstRowIndex, _firstColumnIndex, false, false);
                CellReference crB = new CellReference(_lastRowIndex, _lastColumnIndex, false, false);
                sb.Append(GetType().Name);
                sb.Append(" [").Append(crA.FormatAsString()).Append(':').Append(crB.FormatAsString()).Append("]");
                return sb.ToString();
            }
        }

        private Hashtable _sheetGroupsByBookSheet;

        public FormulaUsedBlankCellSet()
        {
            _sheetGroupsByBookSheet = new Hashtable();
        }

        public void AddCell(int bookIndex, int sheetIndex, int rowIndex, int columnIndex)
        {
            BlankCellSheetGroup sbcg = GetSheetGroup(bookIndex, sheetIndex);
            sbcg.AddCell(rowIndex, columnIndex);
        }

        private BlankCellSheetGroup GetSheetGroup(int bookIndex, int sheetIndex)
        {
            BookSheetKey key = new BookSheetKey(bookIndex, sheetIndex);

            BlankCellSheetGroup result = (BlankCellSheetGroup)_sheetGroupsByBookSheet[key];
            if (result == null)
            {
                result = new BlankCellSheetGroup();
                _sheetGroupsByBookSheet[key]= result;
            }
            return result;
        }

        public bool ContainsCell(BookSheetKey key, int rowIndex, int columnIndex)
        {
            BlankCellSheetGroup bcsg = (BlankCellSheetGroup)_sheetGroupsByBookSheet[key];
            if (bcsg == null)
            {
                return false;
            }
            return bcsg.ContainsCell(rowIndex, columnIndex);
        }
        public bool IsEmpty
        {
            get
            {
                return _sheetGroupsByBookSheet.Count == 0;
            }
        }
    }
}
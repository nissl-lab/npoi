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

namespace NPOI.SS.Util
{
    using System;

    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using System.Globalization;

    /**
     * For POI internal use only
     *
     * @author Josh Micich
     */
    public class SSCellRange<K> : ICellRange<K> where K:ICell
    {

        private int _height;
        private int _width;
        private K[] _flattenedArray;
        private int _firstRow;
        private int _firstColumn;

        private SSCellRange(int firstRow, int firstColumn, int height, int width, K[] flattenedArray)
        {
            _firstRow = firstRow;
            _firstColumn = firstColumn;
            _height = height;
            _width = width;
            _flattenedArray = flattenedArray;
        }

        public static SSCellRange<K> Create(int firstRow, int firstColumn, int height, int width, List<K> flattenedList, Type cellClass)
        {
            int nItems = flattenedList.Count;
            if (height * width != nItems)
            {
                throw new ArgumentException("Array size mismatch.");
            }

            K[] flattenedArray = (K[])Array.CreateInstance(cellClass, nItems);
            flattenedArray=flattenedList.ToArray();
            return new SSCellRange<K>(firstRow, firstColumn, height, width, flattenedArray);
        }



        public K GetCell(int relativeRowIndex, int relativeColumnIndex)
        {
            if (relativeRowIndex < 0 || relativeRowIndex >= _height)
            {
                throw new IndexOutOfRangeException("Specified row " + relativeRowIndex
                        + " is outside the allowable range (0.." + (_height - 1) + ").");
            }
            if (relativeColumnIndex < 0 || relativeColumnIndex >= _width)
            {
                throw new IndexOutOfRangeException("Specified colummn " + relativeColumnIndex
                        + " is outside the allowable range (0.." + (_width - 1) + ").");
            }
            int flatIndex = _width * relativeRowIndex + relativeColumnIndex;
            return _flattenedArray[flatIndex];
        }
        internal class ArrayIterator<D> :IEnumerator<D>
        {

            private D[] _array;
            private int _index;

            public ArrayIterator(D[] array)
            {
                _array = array;
                _index = 0;
            }

            #region IEnumerator<D> Members
            public bool MoveNext()
            {
                return _index < _array.Length;
            }

            public void Remove()
            {
                throw new NotSupportedException("Cannot remove cells from this CellRange.");
            }

            public void Reset()
            { 
            }

            public D Current
            {
                get
                {
                    if (_index >= _array.Length)
                    {
                        throw new ArgumentNullException(_index.ToString(CultureInfo.CurrentCulture));
                    }
                    return _array[_index++];
                }
            }


            #endregion

            #region IDisposable Members

            public void Dispose()
            {
                //do nothing?
            }

            #endregion

            #region IEnumerator Members

            object System.Collections.IEnumerator.Current
            {
                get { return this.Current; }
            }

            #endregion
        }

        #region CellRange<K> Members


        public K TopLeftCell
        {
            get { return _flattenedArray[0]; }
        }

        public K[] FlattenedCells
        {
            get {
                return (K[])_flattenedArray.Clone();
            }
        }

        public K[][] Cells
        {
            get {
                Type itemCls = _flattenedArray.GetType();
                K[][] result = (K[][])Array.CreateInstance(itemCls, _height);
                itemCls = itemCls.GetElementType();
                for (int r = _height - 1; r >= 0; r--)
                {
                    K[] row = (K[])Array.CreateInstance(itemCls, _width);
                    int flatIndex = _width * r;
                    Array.Copy(_flattenedArray, flatIndex, row, 0, _width);
                }
                return result;
            }
        }

        public int Height
        {
            get
            {
                return _height;
            }
        }
        public int Width
        {
            get
            {
                return _width;
            }
        }
        public int Size
        {
            get
            {
                return _height * _width;
            }
        }

        public String ReferenceText
        {
            get
            {
                CellRangeAddress cra = new CellRangeAddress(_firstRow, _firstRow + _height - 1, _firstColumn, _firstColumn + _width - 1);
                return cra.FormatAsString();
            }
        }
        #endregion




        #region IEnumerable<K> Members

        public IEnumerator<K> GetEnumerator()
        {
            return new ArrayIterator<K>(_flattenedArray);
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}




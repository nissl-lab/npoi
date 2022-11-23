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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.Util;
using System;
using System.Collections.Generic;

namespace NPOI.XSSF.UserModel.Helpers
{
    /// <summary>
    /// Helper class for dealing with the Column Settings on a CT_Worksheet 
    /// (the data part of a sheet). Note - within POI, we use 0 based column 
    /// indexes, but the column defInitions in the XML are 1 based!
    /// </summary>
    public class ColumnHelper
    {

        private readonly CT_Worksheet worksheet;

        #region Constructor
        public ColumnHelper(CT_Worksheet worksheet)
        {

            this.worksheet = worksheet;
            CleanColumns();
        }
        #endregion

        #region Public methods
        public void CleanColumns()
        {
            TreeSet<CT_Col> trackedCols = 
                new TreeSet<CT_Col>(CTColComparator.BY_MIN_MAX);
            CT_Cols newCols = new CT_Cols();
            CT_Cols[] colsArray = worksheet.GetColsList().ToArray();
            int i;

            for (i = 0; i < colsArray.Length; i++)
            {
                CT_Cols cols = colsArray[i];
                CT_Col[] colArray = cols.GetColList().ToArray();

                foreach (CT_Col col in colArray)
                {
                    AddCleanColIntoCols(newCols, col, trackedCols);
                }
            }

            for (int y = i - 1; y >= 0; y--)
            {
                worksheet.RemoveCols(y);
            }

            newCols.SetColArray(
                trackedCols.ToArray(new CT_Col[trackedCols.Count]));
            worksheet.AddNewCols();
            worksheet.SetColsArray(0, newCols);
        }

        public CT_Cols AddCleanColIntoCols(CT_Cols cols, CT_Col newCol)
        {
            // Performance issue. If we encapsulated management of min/max in
            // this class then we could keep trackedCols as state, making this
            // log(N) rather than Nlog(N). We do this for the initial
            // read above.
            TreeSet<CT_Col> trackedCols = 
                new TreeSet<CT_Col>(CTColComparator.BY_MIN_MAX);
            trackedCols.AddAll(cols.GetColList());
            AddCleanColIntoCols(cols, newCol, trackedCols);
            cols.SetColArray(trackedCols.ToArray(new CT_Col[0]));
            return cols;
        }

        public static void SortColumns(CT_Cols newCols)
        {
            List<CT_Col> colArray = newCols.GetColList();
            colArray.Sort(new CTColComparator());
            newCols.SetColArray(colArray);
        }

        public CT_Col CloneCol(CT_Cols cols, CT_Col col)
        {
            CT_Col newCol = cols.AddNewCol();
            newCol.min = col.min;
            newCol.max = col.max;
            SetColumnAttributes(col, newCol);
            return newCol;
        }

        /// <summary>
        /// Returns the Column at the given 0 based index
        /// </summary>
        /// <param name="index"></param>
        /// <param name="splitColumns"></param>
        /// <returns></returns>
        public CT_Col GetColumn(long index, bool splitColumns)
        {
            return GetColumn1Based(index + 1, splitColumns);
        }

        /// <summary>
        /// Returns the Column at the given 1 based index. POI default is 0 
        /// based, but the file stores as 1 based.
        /// </summary>
        /// <param name="index1"></param>
        /// <param name="splitColumns"></param>
        /// <returns></returns>
        public CT_Col GetColumn1Based(long index1, bool splitColumns)
        {
            CT_Cols colsArray = worksheet.GetColsArray(0);

            // Fetching the array is quicker than working on the new style
            //  list, assuming we need to read many of them (which we often do),
            //  and assuming we're not making many changes (which we're not)
            CT_Col[] cols = colsArray.GetColList().ToArray();
            for (int i = 0; i < cols.Length; i++)
            {
                CT_Col colArray = cols[i];
                long colMin = colArray.min;
                long colMax = colArray.max;
                if (colMin <= index1 && colMax >= index1)
                {
                    if (splitColumns)
                    {
                        if (colMin < index1)
                        {
                            InsertCol(
                                colsArray, 
                                colMin, 
                                index1 - 1, 
                                new CT_Col[] { colArray });
                        }

                        if (colMax > index1)
                        {
                            InsertCol(
                                colsArray, 
                                index1 + 1, 
                                colMax, 
                                new CT_Col[] { colArray });
                        }

                        colArray.min = (uint)index1;
                        colArray.max = (uint)index1;
                    }

                    return colArray;
                }
            }

            return null;
        }

        /// <summary>
        /// Does the column at the given 0 based index exist in the supplied 
        /// list of column defInitions?
        /// </summary>
        /// <param name="cols"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool ColumnExists(CT_Cols cols, long index)
        {
            return ColumnExists1Based(cols, index + 1);
        }

        public void SetColumnAttributes(CT_Col fromCol, CT_Col toCol)
        {

            if (fromCol.IsSetBestFit())
            {
                toCol.bestFit = fromCol.bestFit;
            }

            if (fromCol.IsSetCustomWidth())
            {
                toCol.customWidth = fromCol.customWidth;
            }

            if (fromCol.IsSetHidden())
            {
                toCol.hidden = fromCol.hidden;
            }

            if (fromCol.IsSetStyle())
            {
                toCol.style = fromCol.style;
            }

            if (fromCol.IsSetWidth())
            {
                toCol.width = fromCol.width;
                toCol.widthSpecified = fromCol.widthSpecified;
            }

            if (fromCol.IsSetCollapsed())
            {
                toCol.collapsed = fromCol.collapsed;
                toCol.collapsedSpecified = fromCol.collapsedSpecified;
            }

            if (fromCol.IsSetPhonetic())
            {
                toCol.phonetic = fromCol.phonetic;
            }

            if (fromCol.IsSetOutlineLevel())
            {
                toCol.outlineLevel = fromCol.outlineLevel;
            }

            if (fromCol.IsSetCollapsed())
            {
                toCol.collapsed = fromCol.collapsed;
            }
        }

        public void SetColBestFit(long index, bool bestFit)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, false);
            col.bestFit = bestFit;
        }
        public void SetCustomWidth(long index, bool width)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.customWidth = width;
        }

        public void SetColWidth(long index, double width)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.width = width;
        }

        public void SetColHidden(long index, bool hidden)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.hidden = hidden;
        }

        public void SetColDefaultStyle(long index, ICellStyle style)
        {
            SetColDefaultStyle(index, style.Index);
        }

        public void SetColDefaultStyle(long index, int styleId)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.style = (uint)styleId;
        }

        /// <summary>
        /// Returns -1 if no column is found for the given index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public int GetColDefaultStyle(long index)
        {
            CT_Col column = GetColumn(index, false);

            if (column != null && column.style != null)
            {
                return (int)column.style;
            }

            return -1;
        }

        public int GetIndexOfColumn(CT_Cols cols, CT_Col col)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == col.min 
                    && cols.GetColArray(i).max == col.max)
                {
                    return i;
                }
            }

            return -1;
        }
        #endregion

        #region Internal methods
        /// <summary>
        /// Return the CT_Col at the given (0 based) column index, creating 
        /// it if required.
        /// </summary>
        /// <param name="index1"></param>
        /// <param name="splitColumns"></param>
        /// <returns></returns>
        internal CT_Col GetOrCreateColumn1Based(long index1, bool splitColumns)
        {
            CT_Col col = GetColumn1Based(index1, splitColumns);
            if (col == null)
            {
                col = worksheet.GetColsArray(0).AddNewCol();
                col.min = (uint)index1;
                col.max = (uint)index1;
            }

            return col;
        }
        #endregion

        #region Private methods
        private void AddCleanColIntoCols(CT_Cols cols, CT_Col newCol, TreeSet<CT_Col> trackedCols)
        {
            List<CT_Col> overlapping = GetOverlappingCols(newCol, trackedCols);
            if (overlapping.Count == 0)
            {
                trackedCols.Add(CloneCol(cols, newCol));
                return;
            }

            trackedCols.RemoveAll(overlapping);
            foreach (CT_Col existing in overlapping)
            {
                // We add up to three columns for each existing one:
                // non-overlap before, overlap, non-overlap after.
                long[] overlap = GetOverlap(newCol, existing);

                CT_Col overlapCol = CloneCol(cols, existing, overlap);
                SetColumnAttributes(newCol, overlapCol);
                trackedCols.Add(overlapCol);

                CT_Col beforeCol = existing.min < newCol.min ? existing
                        : newCol;
                long[] before = new long[] {
                    Math.Min(existing.min, newCol.min),
                    overlap[0] - 1 };
                if (before[0] <= before[1])
                {
                    trackedCols.Add(CloneCol(cols, beforeCol, before));
                }

                CT_Col afterCol = existing.max > newCol.max ? existing
                        : newCol;
                long[] after = new long[] { overlap[1] + 1,
                    Math.Max(existing.max, newCol.max) };
                if (after[0] <= after[1])
                {
                    trackedCols.Add(CloneCol(cols, afterCol, after));
                }
            }
        }

        private CT_Col CloneCol(CT_Cols cols, CT_Col col, long[] newRange)
        {
            CT_Col cloneCol = CloneCol(cols, col);
            cloneCol.min = (uint)newRange[0];
            cloneCol.max = (uint)newRange[1];
            return cloneCol;
        }

        private long[] GetOverlap(CT_Col col1, CT_Col col2)
        {
            return GetOverlappingRange(col1, col2);
        }

        private List<CT_Col> GetOverlappingCols(CT_Col newCol, TreeSet<CT_Col> trackedCols)
        {
            CT_Col lower = trackedCols.Lower(newCol);
            TreeSet<CT_Col> potentiallyOverlapping = 
                lower == null 
                ? trackedCols 
                : trackedCols.TailSet(lower, Overlaps(lower, newCol));
            List<CT_Col> overlapping = new List<CT_Col>();
            foreach (CT_Col existing in potentiallyOverlapping)
            {
                if (Overlaps(newCol, existing))
                {
                    overlapping.Add(existing);
                }
                else
                {
                    break;
                }
            }

            return overlapping;
        }

        private bool Overlaps(CT_Col col1, CT_Col col2)
        {
            return NumericRanges.GetOverlappingType(ToRange(col1), ToRange(col2)) != NumericRanges.NO_OVERLAPS;
        }

        private long[] GetOverlappingRange(CT_Col col1, CT_Col col2)
        {
            return NumericRanges.GetOverlappingRange(ToRange(col1), ToRange(col2));
        }

        private long[] ToRange(CT_Col col)
        {
            return new long[] { col.min, col.max };
        }

        /// <summary>
        /// Insert a new CT_Col at position 0 into cols, Setting min=min, 
        /// max=max and copying all the colsWithAttributes array cols 
        /// attributes into newCol
        /// </summary>
        /// <param name="cols"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <param name="colsWithAttributes"></param>
        /// <returns></returns>
        private CT_Col InsertCol(CT_Cols cols, long min, long max,
            CT_Col[] colsWithAttributes)
        {
            return InsertCol(cols, min, max, colsWithAttributes, false, null);
        }

        private CT_Col InsertCol(
            CT_Cols cols, 
            long min, 
            long max,
            CT_Col[] colsWithAttributes, 
            bool ignoreExistsCheck, 
            CT_Col overrideColumn)
        {
            if (ignoreExistsCheck || !ColumnExists(cols, min, max))
            {
                CT_Col newCol = cols.InsertNewCol(0);
                newCol.min = (uint)min;
                newCol.max = (uint)max;
                foreach (CT_Col col in colsWithAttributes)
                {
                    SetColumnAttributes(col, newCol);
                }

                if (overrideColumn != null)
                {
                    SetColumnAttributes(overrideColumn, newCol);
                }

                return newCol;
            }

            return null;
        }

        private bool ColumnExists1Based(CT_Cols cols, long index1)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == index1)
                {
                    return true;
                }
            }

            return false;
        }

        private bool ColumnExists(CT_Cols cols, long min, long max)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == min 
                    && cols.GetColArray(i).max == max)
                {
                    return true;
                }
            }

            return false;
        }
        #endregion

        #region TreeSet class
        public class TreeSet<T>
        {
            private readonly SortedList<T, object> innerObj;
            private readonly IComparer<T> comparer;

            public TreeSet(IComparer<T> comparer)
            {
                this.comparer = comparer;
                innerObj = new SortedList<T, object>(comparer);
            }

            public T First()
            {
                IEnumerator<T> enumerator = innerObj.Keys.GetEnumerator();
                if (enumerator.MoveNext())
                {
                    return enumerator.Current;
                }

                return default;
            }

            public T Higher(T element)
            {
                IEnumerator<T> enumerator = innerObj.Keys.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (innerObj.Comparer.Compare(enumerator.Current, element) > 0)
                    {
                        return enumerator.Current;
                    }
                }

                return default;
            }

            public void Add(T item)
            {
                if (!innerObj.ContainsKey(item))
                {
                    innerObj.Add(item, null);
                }
            }

            public bool Remove(T item)
            {
                return innerObj.Remove(item);
            }

            public int Count
            {
                get { return innerObj.Count; }
            }

            public void CopyTo(T[] target)
            {
                for (int i = 0; i < innerObj.Count; i++)
                {
                    target[i] = innerObj.Keys[i];
                }
            }

            public IEnumerator<T> GetEnumerator()
            {
                return innerObj.Keys.GetEnumerator();
            }

            public T[] ToArray(T[] a)
            {
                List<T> ts = new List<T>();
                ts.AddRange(innerObj.Keys);
                if (a.Length < Count)
                {
                    // Make a new array of a's runtime type, but my contents:
                    return ts.ToArray();
                }

                Array.Copy(ts.ToArray(), 0, a, 0, Count);
                if (a.Length > Count)
                {
                    a[Count] = default;
                }

                return a;
            }

            internal void AddAll(List<T> list)
            {
                foreach (T item in list)
                {
                    if (!innerObj.ContainsKey(item))
                    {
                        innerObj.Add(item, null);
                    }
                }
            }

            internal void RemoveAll(List<T> list)
            {
                foreach (T item in list)
                {
                    innerObj.Remove(item);
                }
            }

            internal T Lower(T element)
            {
                IEnumerator<T> enumerator = innerObj.Keys.GetEnumerator();
                T prevElement = default;
                while (enumerator.MoveNext())
                {
                    if (innerObj.Comparer.Compare(enumerator.Current, element) >= 0)
                    {
                        return prevElement;
                    }

                    prevElement = enumerator.Current;
                }

                return prevElement;
            }

            /// <summary>
            /// Returns a view of the portion of this map whose keys are 
            /// greater than (or equal to, if inclusive is true) fromKey. 
            /// </summary>
            /// <param name="fromElement"></param>
            /// <param name="inclusive"></param>
            /// <returns></returns>
            internal TreeSet<T> TailSet(T fromElement, bool inclusive)
            {
                TreeSet<T> set = new TreeSet<T>(comparer);

                IEnumerator<T> enumerator = innerObj.Keys.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (inclusive)
                    {
                        if (innerObj.Comparer.Compare(enumerator.Current, fromElement) >= 0)
                        {
                            set.Add(enumerator.Current);
                        }
                    }
                    else
                    {
                        if (innerObj.Comparer.Compare(enumerator.Current, fromElement) > 0)
                        {
                            set.Add(enumerator.Current);
                        }
                    }
                }

                return set;
            }
        }
        #endregion
    }
}

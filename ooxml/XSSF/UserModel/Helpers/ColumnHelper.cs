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

using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Util;
namespace NPOI.XSSF.UserModel.Helpers
{
    /**
     * Helper class for dealing with the Column Settings on
     *  a CT_Worksheet (the data part of a sheet).
     * Note - within POI, we use 0 based column indexes, but
     *  the column defInitions in the XML are 1 based!
     */
    public class ColumnHelper
    {

        private CT_Worksheet worksheet;
        //private CT_Cols newCols;

        public ColumnHelper(CT_Worksheet worksheet)
        {

            this.worksheet = worksheet;
            CleanColumns();
        }
        public void CleanColumns()
        {
            TreeSet<CT_Col> trackedCols = new TreeSet<CT_Col>(CTColComparator.BY_MIN_MAX);
            CT_Cols newCols = new CT_Cols();
            CT_Cols[] colsArray = worksheet.GetColsList().ToArray();
            int i = 0;
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

            newCols.SetColArray(trackedCols.ToArray(new CT_Col[trackedCols.Count]));
            worksheet.AddNewCols();
            worksheet.SetColsArray(0, newCols);
            //this.newCols = new CT_Cols();

            //CT_Cols aggregateCols = new CT_Cols();
            //List<CT_Cols> colsList = worksheet.GetColsList();
            //if (colsList == null)
            //{
            //    return;
            //}

            //foreach (CT_Cols cols in colsList)
            //{
            //    foreach (CT_Col col in cols.GetColList())
            //    {
            //        CloneCol(aggregateCols, col);
            //    }
            //}

            //SortColumns(aggregateCols);

            //CT_Col[] colArray = aggregateCols.GetColList().ToArray();
            //SweepCleanColumns(newCols, colArray, null);

            //int i = colsList.Count;
            //for (int y = i - 1; y >= 0; y--)
            //{
            //    worksheet.RemoveCols(y);
            //}
            //worksheet.AddNewCols();
            //worksheet.SetColsArray(0, newCols);
        }

        public CT_Cols AddCleanColIntoCols(CT_Cols cols, CT_Col newCol)
        {
            // Performance issue. If we encapsulated management of min/max in this
            // class then we could keep trackedCols as state,
            // making this log(N) rather than Nlog(N). We do this for the initial
            // read above.
            TreeSet<CT_Col> trackedCols = new TreeSet<CT_Col>(CTColComparator.BY_MIN_MAX);
            trackedCols.AddAll(cols.GetColList());
            AddCleanColIntoCols(cols, newCol, trackedCols);
            cols.SetColArray(trackedCols.ToArray(new CT_Col[0]));
            return cols;
        }

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
                // We add up to three columns for each existing one: non-overlap
                // before, overlap, non-overlap after.
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
            cloneCol.min = (uint)(newRange[0]);
            cloneCol.max = (uint)(newRange[1]);
            return cloneCol;
        }

        private long[] GetOverlap(CT_Col col1, CT_Col col2)
        {
            return GetOverlappingRange(col1, col2);
        }

        private List<CT_Col> GetOverlappingCols(CT_Col newCol, TreeSet<CT_Col> trackedCols)
        {
            CT_Col lower = trackedCols.Lower(newCol);
            TreeSet<CT_Col> potentiallyOverlapping = lower == null ? trackedCols : trackedCols.TailSet(lower, Overlaps(lower, newCol));
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

        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public static void SortColumns(CT_Cols newCols)
        {
            List<CT_Col> colArray = newCols.GetColList();
            colArray.Sort(new CTColComparator());
            newCols.SetColArray(colArray);
        }

        public CT_Col CloneCol(CT_Cols cols, CT_Col col)
        {
            CT_Col newCol = cols.AddNewCol();
            newCol.min = (uint)(col.min);
            newCol.max = (uint)(col.max);
            SetColumnAttributes(col, newCol);
            return newCol;
        }

        /**
         * Returns the Column at the given 0 based index
         */
        public CT_Col GetColumn(long index, bool splitColumns)
        {
            return GetColumn1Based(index + 1, splitColumns);
        }
        /**
         * Returns the Column at the given 1 based index.
         * POI default is 0 based, but the file stores
         *  as 1 based.
         */
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
                            InsertCol(colsArray, colMin, (index1 - 1), new CT_Col[] { colArray });
                        }
                        if (colMax > index1)
                        {
                            InsertCol(colsArray, (index1 + 1), colMax, new CT_Col[] { colArray });
                        }
                        colArray.min = (uint)(index1);
                        colArray.max = (uint)(index1);
                    }
                    return colArray;
                }
            }
            return null;
        }

        public class TreeSet<T>
        {
            private SortedList<T, object> innerObj;
            private IComparer<T> comparer;
            public TreeSet(IComparer<T> comparer)
            {
                this.comparer = comparer;
                innerObj = new SortedList<T, object>(comparer);
            }
            public T First()
            {
                IEnumerator<T> enumerator = this.innerObj.Keys.GetEnumerator();
                if (enumerator.MoveNext())
                    return enumerator.Current;
                return default(T);
            }
            public T Higher(T element)
            {
                IEnumerator<T> enumerator = this.innerObj.Keys.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (this.innerObj.Comparer.Compare(enumerator.Current, element) > 0)
                        return enumerator.Current;
                }
                return default(T);
            }
            public void Add(T item)
            {
                if(!this.innerObj.ContainsKey(item))
                    this.innerObj.Add(item,null);
            }
            public bool Remove(T item)
            {
                return this.innerObj.Remove(item);
            }
            public int Count
            {
                get { return this.innerObj.Count; }
            }
            public void CopyTo(T[] target)
            { 
                for (int i = 0; i < this.innerObj.Count; i++)
                {
                    target[i] = (T)this.innerObj.Keys[i];
                }
            }

            public IEnumerator<T> GetEnumerator()
            {
                return this.innerObj.Keys.GetEnumerator();
            }

            public T[] ToArray(T[] a)
            {
                List<T> ts = new List<T>();
                ts.AddRange(this.innerObj.Keys);
                if (a.Length < Count)
                {
                    // Make a new array of a's runtime type, but my contents:
                    return ts.ToArray();
                }

                Array.Copy(ts.ToArray(), 0, a, 0, Count);
                if (a.Length > Count)
                    a[Count] = default(T);
                return a;
            }

            internal void AddAll(List<T> list)
            {
                foreach(var item in list)
                {
                    if (!this.innerObj.ContainsKey(item))
                        this.innerObj.Add(item, null);
                }
            }

            internal void RemoveAll(List<T> list)
            {
                foreach (var item in list)
                    this.innerObj.Remove(item);
            }

            internal T Lower(T element)
            {
                IEnumerator<T> enumerator = this.innerObj.Keys.GetEnumerator();
                T prevElement = default(T);
                while (enumerator.MoveNext())
                {
                    if (this.innerObj.Comparer.Compare(enumerator.Current, element) >= 0)
                    {
                        return prevElement;
                    }
                    prevElement = enumerator.Current;
                }
                return prevElement;
            }

            /// <summary>
            /// Returns a view of the portion of this map whose keys are greater than (or
            /// equal to, if inclusive is true) fromKey. 
            /// </summary>
            /// <param name="fromElement"></param>
            /// <param name="inclusive"></param>
            /// <returns></returns>
            internal TreeSet<T> TailSet(T fromElement, bool inclusive)
            {
                TreeSet<T> set = new TreeSet<T>(comparer);

                IEnumerator<T> enumerator = this.innerObj.Keys.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (inclusive)
                    {
                        if (this.innerObj.Comparer.Compare(enumerator.Current, fromElement) >= 0)
                        {
                            set.Add(enumerator.Current);
                        }
                    }
                    else
                    {
                        if (this.innerObj.Comparer.Compare(enumerator.Current, fromElement) > 0)
                        {
                            set.Add(enumerator.Current);
                        }
                    }
                }

                return set;
            }
        }
        /**
         * @see <a href="http://en.wikipedia.org/wiki/Sweep_line_algorithm">Sweep line algorithm</a>
         */
        private void SweepCleanColumns(CT_Cols cols, CT_Col[] flattenedColsArray, CT_Col overrideColumn)
        {
            List<CT_Col> flattenedCols = new List<CT_Col>(flattenedColsArray);
            TreeSet<CT_Col> currentElements = new TreeSet<CT_Col>(CTColComparator.BY_MAX);
            IEnumerator<CT_Col> flIter = flattenedCols.GetEnumerator();
            CT_Col haveOverrideColumn = null;
            long lastMaxIndex = 0;
            long currentMax = 0;
            IList<CT_Col> toRemove = new List<CT_Col>();
            int pos = -1;
            //while (flIter.hasNext())
            while ((pos + 1) < flattenedCols.Count)
            {
                //CTCol col = flIter.next();
                pos++;
                CT_Col col = flattenedCols[pos]; 
                
                long currentIndex = col.min;
                long colMax = col.max;
                long nextIndex = (colMax > currentMax) ? colMax : currentMax;
                //if (flIter.hasNext()) {
                if((pos+1)<flattenedCols.Count)
                {
                    //nextIndex = flIter.next().getMin();
                    nextIndex = flattenedCols[pos + 1].min;
                    //flIter.previous();
                }
                IEnumerator<CT_Col> iter = currentElements.GetEnumerator();
                toRemove.Clear();
                while (iter.MoveNext())
                {
                    CT_Col elem = iter.Current;
                    if (currentIndex <= elem.max) 
                        break; // all passed elements have been purged
                    
                    toRemove.Add(elem);
                }

                foreach (CT_Col rc in toRemove)
                {
                    currentElements.Remove(rc);
                }
                
                if (!(currentElements.Count == 0) && lastMaxIndex < currentIndex)
                {
                    // we need to process previous elements first
                    CT_Col[] copyCols = new CT_Col[currentElements.Count];
                    currentElements.CopyTo(copyCols);
                    InsertCol(cols, lastMaxIndex, currentIndex - 1, copyCols, true, haveOverrideColumn);
                }
                currentElements.Add(col);
                if (colMax > currentMax) currentMax = colMax;
                if (col.Equals(overrideColumn)) haveOverrideColumn = overrideColumn;
                while (currentIndex <= nextIndex && !(currentElements.Count == 0))
                {
                    NPOI.Util.Collections.HashSet<CT_Col> currentIndexElements = new NPOI.Util.Collections.HashSet<CT_Col>();
                    long currentElemIndex;

                    {
                        // narrow scope of currentElem
                        CT_Col currentElem = currentElements.First();
                        currentElemIndex = currentElem.max;
                        currentIndexElements.Add(currentElem);

                        while (true)
                        {
                            CT_Col higherElem = currentElements.Higher(currentElem);
                            if (higherElem == null || higherElem.max != currentElemIndex)
                                break;
                            currentElem = higherElem;
                            currentIndexElements.Add(currentElem);
                            if (colMax > currentMax) currentMax = colMax;
                            if (col.Equals(overrideColumn)) haveOverrideColumn = overrideColumn;
                        }
                    }

                    //if (currentElemIndex < nextIndex || !flIter.hasNext()) {
                    if (currentElemIndex < nextIndex || !((pos + 1) < flattenedCols.Count))
                    {
                        CT_Col[] copyCols = new CT_Col[currentElements.Count];
                        currentElements.CopyTo(copyCols);
                        InsertCol(cols, currentIndex, currentElemIndex, copyCols, true, haveOverrideColumn);
                        //if (flIter.hasNext()) {
                        if ((pos + 1) < flattenedCols.Count)
                        {
                            if (nextIndex > currentElemIndex)
                            {
                                //currentElements.removeAll(currentIndexElements);
                                foreach (CT_Col rc in currentIndexElements)
                                    currentElements.Remove(rc);
                                if (currentIndexElements.Contains(overrideColumn)) haveOverrideColumn = null;
                            }
                        }
                        else
                        {
                            //currentElements.removeAll(currentIndexElements);
                            foreach (CT_Col rc in currentIndexElements)
                                currentElements.Remove(rc);
                            if (currentIndexElements.Contains(overrideColumn)) haveOverrideColumn = null;
                        }
                        lastMaxIndex = currentIndex = currentElemIndex + 1;
                    }
                    else
                    {
                        lastMaxIndex = currentIndex;
                        currentIndex = nextIndex + 1;
                    }

                }
            }
            SortColumns(cols);
        }

        //public CT_Cols AddCleanColIntoCols(CT_Cols cols, CT_Col col)
        //{
        //    bool colOverlaps = false;
        //    // a Map to remember overlapping columns
        //    Dictionary<long, bool> overlappingCols = new Dictionary<long, bool>();
        //    int sizeOfColArray = cols.sizeOfColArray();
        //    for (int i = 0; i < sizeOfColArray; i++)
        //    {
        //        CT_Col ithCol = cols.GetColArray(i);
        //        long[] range1 = { ithCol.min, ithCol.max };
        //        long[] range2 = { col.min, col.max };
        //        long[] overlappingRange = NumericRanges.GetOverlappingRange(range1,
        //                range2);
        //        int overlappingType = NumericRanges.GetOverlappingType(range1,
        //                range2);
        //        // different behavior required for each of the 4 different
        //        // overlapping types
        //        if (overlappingType == NumericRanges.OVERLAPS_1_MINOR)
        //        {
        //            // move the max border of the ithCol 
        //            // and insert a new column within the overlappingRange with merged column attributes
        //            ithCol.max = (uint)(overlappingRange[0] - 1);
        //            insertCol(cols, overlappingRange[0],
        //                    overlappingRange[1], new CT_Col[] { ithCol, col });
        //            i++;
        //            //CT_Col newCol = insertCol(cols, (overlappingRange[1] + 1), col
        //            //        .max, new CT_Col[] { col });
        //            //i++;
        //        }
        //        else if (overlappingType == NumericRanges.OVERLAPS_2_MINOR)
        //        {
        //            // move the min border of the ithCol 
        //            // and insert a new column within the overlappingRange with merged column attributes
        //            ithCol.min = (uint)(overlappingRange[1] + 1);
        //            insertCol(cols, overlappingRange[0],
        //                    overlappingRange[1], new CT_Col[] { ithCol, col });
        //            i++;
        //            //CT_Col newCol = insertCol(cols, col.min,
        //            //        (overlappingRange[0] - 1), new CT_Col[] { col });
        //            //i++;
        //        }
        //        else if (overlappingType == NumericRanges.OVERLAPS_2_WRAPS)
        //        {
        //            // merge column attributes, no new column is needed
        //            SetColumnAttributes(col, ithCol);
                    
        //        }
        //        else if (overlappingType == NumericRanges.OVERLAPS_1_WRAPS)
        //        {
        //            // split the ithCol in three columns: before the overlappingRange, overlappingRange, and after the overlappingRange
        //            // before overlappingRange
        //            if (col.min != ithCol.min)
        //            {
        //                insertCol(cols, ithCol.min, (col
        //                        .min - 1), new CT_Col[] { ithCol });
        //                i++;
        //            }
        //            // after the overlappingRange
        //            if (col.max != ithCol.max)
        //            {
        //                insertCol(cols, (col.max + 1),
        //                        ithCol.max, new CT_Col[] { ithCol });
        //                i++;
        //            }
        //            // within the overlappingRange
        //            ithCol.min = (uint)(overlappingRange[0]);
        //            ithCol.max = (uint)(overlappingRange[1]);
        //            SetColumnAttributes(col, ithCol);
        //        }
        //        if (overlappingType != NumericRanges.NO_OVERLAPS)
        //        {
        //            colOverlaps = true;
        //            // remember overlapped columns
        //            for (long j = overlappingRange[0]; j <= overlappingRange[1]; j++)
        //            {
        //                overlappingCols.Add(j, true);
        //            }
        //        }
        //    }
        //    if (!colOverlaps)
        //    {
        //        CloneCol(cols, col);
        //    }
        //    else
        //    {
        //        // insert new columns for ranges without overlaps
        //        long colMin = -1;
        //        for (long j = col.min; j <= col.max; j++)
        //        {
        //            if (overlappingCols.ContainsKey(j) && !overlappingCols[j])
        //            {
        //                if (colMin < 0)
        //                {
        //                    colMin = j;
        //                }
        //                if ((j + 1) > col.max || overlappingCols[(j + 1)])
        //                {
        //                    insertCol(cols, colMin, j, new CT_Col[] { col });
        //                    colMin = -1;
        //                }
        //            }
        //        }
        //    }
        //    SortColumns(cols);
        //    return cols;
        //}

        /*
         * Insert a new CT_Col at position 0 into cols, Setting min=min, max=max and
         * copying all the colsWithAttributes array cols attributes into newCol
         */
        private CT_Col InsertCol(CT_Cols cols, long min, long max,
            CT_Col[] colsWithAttributes)
        {
            return InsertCol(cols, min, max, colsWithAttributes, false, null);
        }
        private CT_Col InsertCol(CT_Cols cols, long min, long max,
        CT_Col[] colsWithAttributes, bool ignoreExistsCheck, CT_Col overrideColumn)
        {
            if (ignoreExistsCheck || !ColumnExists(cols, min, max))
            {
                CT_Col newCol = cols.InsertNewCol(0);
                newCol.min = (uint)(min);
                newCol.max = (uint)(max);
                foreach (CT_Col col in colsWithAttributes)
                {
                    SetColumnAttributes(col, newCol);
                }
                if (overrideColumn != null) SetColumnAttributes(overrideColumn, newCol);
                return newCol;
            }
            return null;
        }
        /**
         * Does the column at the given 0 based index exist
         *  in the supplied list of column defInitions?
         */
        public bool ColumnExists(CT_Cols cols, long index)
        {
            return ColumnExists1Based(cols, index + 1);
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

        public void SetColumnAttributes(CT_Col fromCol, CT_Col toCol)
        {

            if (fromCol.IsSetBestFit())
            {
                toCol.bestFit = (fromCol.bestFit);
            }
            if (fromCol.IsSetCustomWidth())
            {
                toCol.customWidth = (fromCol.customWidth);
            }
            if (fromCol.IsSetHidden()) 
            {
                toCol.hidden = (fromCol.hidden);
            }
            if (fromCol.IsSetStyle())
            {
                toCol.style = (fromCol.style);
            }
            if (fromCol.IsSetWidth())
            {
                toCol.width = (fromCol.width);
                toCol.widthSpecified = fromCol.widthSpecified;
            }
            if (fromCol.IsSetCollapsed())
            {
                toCol.collapsed = (fromCol.collapsed);
                toCol.collapsedSpecified = fromCol.collapsedSpecified;
            }
            if (fromCol.IsSetPhonetic())
            {
                toCol.phonetic = (fromCol.phonetic);
            }
            if (fromCol.IsSetOutlineLevel())
            {
                toCol.outlineLevel = (fromCol.outlineLevel);
            }
            if (fromCol.IsSetCollapsed())
            {
                toCol.collapsed = fromCol.collapsed;
            }
        }

        public void SetColBestFit(long index, bool bestFit)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, false);
            col.bestFit = (bestFit);
        }
        public void SetCustomWidth(long index, bool width)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.customWidth = (width);
        }

        public void SetColWidth(long index, double width)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.width = (width);
        }

        public void SetColHidden(long index, bool hidden)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.hidden = (hidden);
        }

        /**
         * Return the CT_Col at the given (0 based) column index,
         *  creating it if required.
         */
        internal CT_Col GetOrCreateColumn1Based(long index1, bool splitColumns)
        {
            CT_Col col = GetColumn1Based(index1, splitColumns);
            if (col == null)
            {
                col = worksheet.GetColsArray(0).AddNewCol();
                col.min = (uint)(index1);
                col.max = (uint)(index1);
            }
            return col;
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

        // Returns -1 if no column is found for the given index
        public int GetColDefaultStyle(long index)
        {
            if (GetColumn(index, false) != null)
            {
                return (int)GetColumn(index, false).style;
            }
            return -1;
        }

        private bool ColumnExists(CT_Cols cols, long min, long max)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == min && cols.GetColArray(i).max == max)
                {
                    return true;
                }
            }
            return false;
        }

        public int GetIndexOfColumn(CT_Cols cols, CT_Col col)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == col.min && cols.GetColArray(i).max == col.max)
                {
                    return i;
                }
            }
            return -1;
        }
    }

}

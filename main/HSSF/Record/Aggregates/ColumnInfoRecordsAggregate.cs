/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/

namespace NPOI.HSSF.Record.Aggregates
{

    using System;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Model;
    using System.Globalization;


    /// <summary>
    /// @author Glen Stampoultzis
    /// </summary>
    public class ColumnInfoRecordsAggregate : RecordAggregate
    {
        private class CIRComparator : IComparer
        {
            public static IComparer instance = new CIRComparator();
            private CIRComparator()
            {
                // enforce singleton
            }
            public int Compare(Object a, Object b)
            {
                return CompareColInfos((ColumnInfoRecord)a, (ColumnInfoRecord)b);
            }
            public static int CompareColInfos(ColumnInfoRecord a, ColumnInfoRecord b)
            {
                return a.FirstColumn - b.FirstColumn;
            }
        }

        //    int     size     = 0;
        ArrayList records = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnInfoRecordsAggregate"/> class.
        /// </summary>
        public ColumnInfoRecordsAggregate()
        {
            records = new ArrayList();
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnInfoRecordsAggregate"/> class.
        /// </summary>
        /// <param name="rs">The rs.</param>
        public ColumnInfoRecordsAggregate(RecordStream rs): this()
        {
            bool isInOrder = true;
            ColumnInfoRecord cirPrev = null;
            while (rs.PeekNextClass() == typeof(ColumnInfoRecord))
            {
                ColumnInfoRecord cir = (ColumnInfoRecord)rs.GetNext();
                records.Add(cir);
                if (cirPrev != null && CIRComparator.CompareColInfos(cirPrev, cir) > 0)
                {
                    isInOrder = false;
                }
                cirPrev = cir;
            }
            if (records.Count < 1)
            {
                throw new InvalidOperationException("No column info records found");
            }
            if (!isInOrder)
            {
                records.Sort(CIRComparator.instance);
            }
        }
        /** It's an aggregate... just made something up */
        public override short Sid
        {
            get { return -1012; }
        }
        /// <summary>
        /// Gets the num columns.
        /// </summary>
        /// <value>The num columns.</value>
        public int NumColumns
        {
            get
            {
                return records.Count;
            }
        }
        /// <summary>
        /// Gets the size of the record.
        /// </summary>
        /// <value>The size of the record.</value>
        public override int RecordSize
        {
            get
            {
                int size = 0;
                for (IEnumerator iterator = records.GetEnumerator(); iterator.MoveNext(); )
                    size += ((ColumnInfoRecord)iterator.Current).RecordSize;
                return size;
            }
        }

        public IEnumerator GetEnumerator()
        {
            return records.GetEnumerator();
        }

        /**
         * Performs a deep Clone of the record
         */
        public Object Clone()
        {
            ColumnInfoRecordsAggregate rec = new ColumnInfoRecordsAggregate();
            for (int k = 0; k < records.Count; k++)
            {
                ColumnInfoRecord ci = (ColumnInfoRecord)records[k];
                ci = (ColumnInfoRecord)ci.Clone();
                rec.records.Add(ci);    
            }
            return rec;
        }

        /// <summary>
        /// Inserts a column into the aggregate (at the end of the list).
        /// </summary>
        /// <param name="col">The column.</param>
        public void InsertColumn(ColumnInfoRecord col)
        {
            records.Add(col);
            records.Sort(CIRComparator.instance);
        }

        /// <summary>
        /// Inserts a column into the aggregate (at the position specified
        /// by index
        /// </summary>
        /// <param name="idx">The index.</param>
        /// <param name="col">The columninfo.</param>
        public void InsertColumn(int idx, ColumnInfoRecord col)
        {
            records.Insert(idx, col);
        }

        /// <summary>
        /// called by the class that is responsible for writing this sucker.
        /// Subclasses should implement this so that their data is passed back in a
        /// byte array.
        /// </summary>
        /// <param name="offset">offset to begin writing at</param>
        /// <param name="data">byte array containing instance data</param>
        /// <returns>number of bytes written</returns>
        public override int Serialize(int offset, byte[] data)
        {
            IEnumerator itr = records.GetEnumerator();
            int pos = offset;

            while (itr.MoveNext())
            {
                pos += ((Record)itr.Current).Serialize(pos, data);
            }
            return pos - offset;
        }
        /// <summary>
        /// Visit each of the atomic BIFF records contained in this {@link RecordAggregate} in the order
        /// that they should be written to file.  Implementors may or may not return the actual
        /// Records being used to manage POI's internal implementation.  Callers should not
        /// assume either way, and therefore only attempt to modify those Records after cloning
        /// </summary>
        /// <param name="rv"></param>
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            int nItems = records.Count;
            if (nItems < 1)
            {
                return;
            }
            ColumnInfoRecord cirPrev = null;
            for (int i = 0; i < nItems; i++)
            {
                ColumnInfoRecord cir = (ColumnInfoRecord)records[i];
                rv.VisitRecord(cir);
                if (cirPrev != null && CIRComparator.CompareColInfos(cirPrev, cir) > 0)
                {
                    // Excel probably wouldn't mind, but there is much logic in this class
                    // that assumes the column info records are kept in order
                    throw new InvalidOperationException("Column info records are out of order");
                }
                cirPrev = cir;
            }
        }
        /// <summary>
        /// Finds the start of column outline group.
        /// </summary>
        /// <param name="idx">The idx.</param>
        /// <returns></returns>
        public int FindStartOfColumnOutlineGroup(int idx)
        {
            // Find the start of the Group.
            ColumnInfoRecord columnInfo = (ColumnInfoRecord)records[idx];
            int level = columnInfo.OutlineLevel;
            while (idx != 0)
            {
                ColumnInfoRecord prevColumnInfo = (ColumnInfoRecord)records[idx - 1];
                if (columnInfo.FirstColumn - 1 == prevColumnInfo.LastColumn)
                {
                    if (prevColumnInfo.OutlineLevel < level)
                    {
                        break;
                    }
                    idx--;
                    columnInfo = prevColumnInfo;
                }
                else
                {
                    break;
                }
            }

            return idx;
        }

        /// <summary>
        /// Finds the end of column outline group.
        /// </summary>
        /// <param name="idx">The idx.</param>
        /// <returns></returns>
        public int FindEndOfColumnOutlineGroup(int idx)
        {
            // Find the end of the Group.
            ColumnInfoRecord columnInfo = (ColumnInfoRecord)records[idx];
            int level = columnInfo.OutlineLevel;
            while (idx < records.Count - 1)
            {
                ColumnInfoRecord nextColumnInfo = (ColumnInfoRecord)records[idx + 1];
                if (columnInfo.LastColumn + 1 == nextColumnInfo.FirstColumn)
                {
                    if (nextColumnInfo.OutlineLevel < level)
                    {
                        break;
                    }
                    idx++;
                    columnInfo = nextColumnInfo;
                }
                else
                {
                    break;
                }
            }

            return idx;
        }

        /// <summary>
        /// Gets the col info.
        /// </summary>
        /// <param name="idx">The idx.</param>
        /// <returns></returns>
        public ColumnInfoRecord GetColInfo(int idx)
        {
            return (ColumnInfoRecord)records[idx];
        }

        /// <summary>
        /// Determines whether [is column group collapsed] [the specified idx].
        /// </summary>
        /// <param name="idx">The idx.</param>
        /// <returns>
        /// 	<c>true</c> if [is column group collapsed] [the specified idx]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsColumnGroupCollapsed(int idx)
        {
            int endOfOutlineGroupIdx = FindEndOfColumnOutlineGroup(idx);
            int nextColInfoIx = endOfOutlineGroupIdx + 1;
            if (nextColInfoIx >= records.Count)
            {
                return false;
            }
            ColumnInfoRecord nextColInfo = GetColInfo(nextColInfoIx);
            if (!GetColInfo(endOfOutlineGroupIdx).IsAdjacentBefore(nextColInfo))
            {
                return false;
            }
            return nextColInfo.IsCollapsed;
        }


        /// <summary>
        /// Determines whether [is column group hidden by parent] [the specified idx].
        /// </summary>
        /// <param name="idx">The idx.</param>
        /// <returns>
        /// 	<c>true</c> if [is column group hidden by parent] [the specified idx]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsColumnGroupHiddenByParent(int idx)
        {
            // Look out outline details of end
            int endLevel = 0;
            bool endHidden = false;
            int endOfOutlineGroupIdx = FindEndOfColumnOutlineGroup(idx);
            if (endOfOutlineGroupIdx < records.Count)
            {
                ColumnInfoRecord nextInfo = GetColInfo(endOfOutlineGroupIdx + 1);
                if (GetColInfo(endOfOutlineGroupIdx).IsAdjacentBefore(nextInfo))
                {
                    endLevel = nextInfo.OutlineLevel;
                    endHidden = nextInfo.IsHidden;
                }
            }
            // Look out outline details of start
            int startLevel = 0;
            bool startHidden = false;
            int startOfOutlineGroupIdx = FindStartOfColumnOutlineGroup(idx);
            if (startOfOutlineGroupIdx > 0)
            {
                ColumnInfoRecord prevInfo = GetColInfo(startOfOutlineGroupIdx - 1);
                if (prevInfo.IsAdjacentBefore(GetColInfo(startOfOutlineGroupIdx)))
                {
                    startLevel = prevInfo.OutlineLevel;
                    startHidden = prevInfo.IsHidden;
                }
            }
            if (endLevel > startLevel)
            {
                return endHidden;
            }
            return startHidden;
        }

        /// <summary>
        /// Collapses the column.
        /// </summary>
        /// <param name="columnNumber">The column number.</param>
        public void CollapseColumn(int columnNumber)
        {
            int idx = FindColInfoIdx(columnNumber, 0);
            if (idx == -1)
                return;

            // Find the start of the group.
            int groupStartColInfoIx = FindStartOfColumnOutlineGroup(idx);
            ColumnInfoRecord columnInfo = GetColInfo(groupStartColInfoIx);

            // Hide all the columns until the end of the group
            int lastColIx = SetGroupHidden(groupStartColInfoIx, columnInfo.OutlineLevel, true);

            // Write collapse field
            SetColumn(lastColIx + 1, null, null, null, null, true);
        }

        /// <summary>
        /// Expands the column.
        /// </summary>
        /// <param name="columnNumber">The column number.</param>
        public void ExpandColumn(int columnNumber)
        {
            int idx = FindColInfoIdx(columnNumber, 0);
            if (idx == -1)
                return;

            // If it is already exapanded do nothing.
            if (!IsColumnGroupCollapsed(idx))
                return;

            // Find the start of the Group.
            int startIdx = FindStartOfColumnOutlineGroup(idx);
            ColumnInfoRecord columnInfo = GetColInfo(startIdx);

            // Find the end of the Group.
            int endIdx = FindEndOfColumnOutlineGroup(idx);
            ColumnInfoRecord endColumnInfo = GetColInfo(endIdx);

            // expand:
            // colapsed bit must be UnSet
            // hidden bit Gets UnSet _if_ surrounding Groups are expanded you can determine
            //   this by looking at the hidden bit of the enclosing Group.  You will have
            //   to look at the start and the end of the current Group to determine which
            //   is the enclosing Group
            // hidden bit only is altered for this outline level.  ie.  don't Uncollapse contained Groups
            if (!IsColumnGroupHiddenByParent(idx))
            {
                for (int i = startIdx; i <= endIdx; i++)
                {
                    if (columnInfo.OutlineLevel == GetColInfo(i).OutlineLevel)
                        GetColInfo(i).IsHidden = false;
                }
            }

            // Write collapse field
            SetColumn(columnInfo.LastColumn + 1, null, null, null, null, false);
        }

        /**
         * Sets all non null fields into the <c>ci</c> parameter.
         */
        private static void SetColumnInfoFields(ColumnInfoRecord ci, short? xfStyle, int? width,
                    int? level, Boolean? hidden, Boolean? collapsed)
        {
            if (xfStyle != null)
            {
                ci.XFIndex = Convert.ToInt16(xfStyle, CultureInfo.InvariantCulture);
            }
            if (width != null)
            {
                ci.ColumnWidth = Convert.ToInt32(width, CultureInfo.InvariantCulture);
            }
            if (level != null)
            {
                ci.OutlineLevel = (short)level;
            }
            if (hidden != null)
            {
                ci.IsHidden = Convert.ToBoolean(hidden, CultureInfo.InvariantCulture);
            }
            if (collapsed != null)
            {
                ci.IsCollapsed = Convert.ToBoolean(collapsed, CultureInfo.InvariantCulture);
            }
        }
        /// <summary>
        /// Attempts to merge the col info record at the specified index
        /// with either or both of its neighbours
        /// </summary>
        /// <param name="colInfoIx">The col info ix.</param>
        private void AttemptMergeColInfoRecords(int colInfoIx)
        {
            int nRecords = records.Count;
            if (colInfoIx < 0 || colInfoIx >= nRecords)
            {
                throw new ArgumentException("colInfoIx " + colInfoIx
                        + " is out of range (0.." + (nRecords - 1) + ")");
            }
            ColumnInfoRecord currentCol = GetColInfo(colInfoIx);
            int nextIx = colInfoIx + 1;
            if (nextIx < nRecords)
            {
                if (MergeColInfoRecords(currentCol, GetColInfo(nextIx)))
                {
                    records.RemoveAt(nextIx);
                }
            }
            if (colInfoIx > 0)
            {
                if (MergeColInfoRecords(GetColInfo(colInfoIx - 1), currentCol))
                {
                    records.RemoveAt(colInfoIx);
                }
            }
        }
        /**
    * merges two column info records (if they are adjacent and have the same formatting, etc)
    * @return <c>false</c> if the two column records could not be merged
    */
        private static bool MergeColInfoRecords(ColumnInfoRecord ciA, ColumnInfoRecord ciB)
        {
            if (ciA.IsAdjacentBefore(ciB) && ciA.FormatMatches(ciB))
            {
                ciA.LastColumn = ciB.LastColumn;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Sets all adjacent columns of the same outline level to the specified hidden status.
        /// </summary>
        /// <param name="pIdx">the col info index of the start of the outline group.</param>
        /// <param name="level">The level.</param>
        /// <param name="hidden">The hidden.</param>
        /// <returns>the column index of the last column in the outline group</returns>
        private int SetGroupHidden(int pIdx, int level, bool hidden)
        {
            int idx = pIdx;
            ColumnInfoRecord columnInfo = GetColInfo(idx);
            while (idx < records.Count)
            {
                columnInfo.IsHidden = (hidden);
                if (idx + 1 < records.Count)
                {
                    ColumnInfoRecord nextColumnInfo = GetColInfo(idx + 1);
                    if (!columnInfo.IsAdjacentBefore(nextColumnInfo))
                    {
                        break;
                    }
                    if (nextColumnInfo.OutlineLevel < level)
                    {
                        break;
                    }
                    columnInfo = nextColumnInfo;
                }
                idx++;
            }
            return columnInfo.LastColumn;
        }
        /// <summary>
        /// Sets the column.
        /// </summary>
        /// <param name="targetColumnIx">The target column ix.</param>
        /// <param name="xfIndex">Index of the xf.</param>
        /// <param name="width">The width.</param>
        /// <param name="level">The level.</param>
        /// <param name="hidden">The hidden.</param>
        /// <param name="collapsed">The collapsed.</param>
        public void SetColumn(int targetColumnIx, short? xfIndex, int? width, int? level, bool? hidden, bool? collapsed)
        {
            ColumnInfoRecord ci = null;
            int k = 0;

            for (k = 0; k < records.Count; k++)
            {
                ColumnInfoRecord tci = (ColumnInfoRecord)records[k];
                if (tci.ContainsColumn(targetColumnIx))
                {
                    ci = tci;
                    break;
                }
                if (tci.FirstColumn > targetColumnIx)
                {
                    // call targetColumnIx infos after k are for later targetColumnIxs
                    break; // exit now so k will be the correct insert pos
                }
            }

            if (ci == null)
            {
                // okay so there IsN'T a targetColumnIx info record that cover's this targetColumnIx so lets Create one!
                ColumnInfoRecord nci = new ColumnInfoRecord();

                nci.FirstColumn = targetColumnIx;
                nci.LastColumn = targetColumnIx;
                SetColumnInfoFields(nci, xfIndex, width, level, hidden, collapsed);
                InsertColumn(k, nci);
                AttemptMergeColInfoRecords(k);
                return;
            }

            bool styleChanged = ci.XFIndex != xfIndex;
            bool widthChanged = ci.ColumnWidth != width;
            bool levelChanged = ci.OutlineLevel != level;
            bool hiddenChanged = ci.IsHidden != hidden;
            bool collapsedChanged = ci.IsCollapsed != collapsed;
            bool targetColumnIxChanged = styleChanged || widthChanged || levelChanged || hiddenChanged || collapsedChanged;
            if (!targetColumnIxChanged)
            {
                // do nothing...nothing Changed.
                return;
            }
            if ((ci.FirstColumn == targetColumnIx)
                     && (ci.LastColumn == targetColumnIx))
            {                               // if its only for this cell then
                // ColumnInfo ci for a single column, the target column
                SetColumnInfoFields(ci, xfIndex, width, level, hidden, collapsed);
                AttemptMergeColInfoRecords(k);
                return;
            }
            if ((ci.FirstColumn == targetColumnIx)
                     || (ci.LastColumn == targetColumnIx))
            {
                // The target column is at either end of the multi-column ColumnInfo ci
                // we'll just divide the info and create a new one
                if (ci.FirstColumn == targetColumnIx)
                {
                    ci.FirstColumn = targetColumnIx + 1;
                }
                else
                {
                    ci.LastColumn = targetColumnIx - 1;
                    k++; // adjust insert pos to insert after
                }
                ColumnInfoRecord nci = CopyColInfo(ci);

                nci.FirstColumn = targetColumnIx;
                nci.LastColumn = targetColumnIx;

                SetColumnInfoFields(nci, xfIndex, width, level, hidden, collapsed);

                InsertColumn(k, nci);
                AttemptMergeColInfoRecords(k);
            }
            else
            {
                //split to 3 records
                ColumnInfoRecord ciStart = ci;
                ColumnInfoRecord ciMid = CopyColInfo(ci);
                ColumnInfoRecord ciEnd = CopyColInfo(ci);
                int lastcolumn = ci.LastColumn;

                ciStart.LastColumn = (targetColumnIx - 1);

                ciMid.FirstColumn=(targetColumnIx);
                ciMid.LastColumn=(targetColumnIx);
                SetColumnInfoFields(ciMid, xfIndex, width, level, hidden, collapsed);
                InsertColumn(++k, ciMid);

                ciEnd.FirstColumn = (targetColumnIx + 1);
                ciEnd.LastColumn = (lastcolumn);
                InsertColumn(++k, ciEnd);
                // no need to attemptMergeColInfoRecords because we 
                // know both on each side are different
            }
        }
        private ColumnInfoRecord CopyColInfo(ColumnInfoRecord ci)
        {
            return (ColumnInfoRecord)ci.Clone();
        }

        /**
         * Sets all non null fields into the <c>ci</c> parameter.
         */
        private void SetColumnInfoFields(ColumnInfoRecord ci, short xfStyle, short width, int level, bool hidden, bool collapsed)
        {
            ci.XFIndex = (xfStyle);
            ci.ColumnWidth = (width);
            ci.OutlineLevel = (short)level;
            ci.IsHidden = (hidden);
            ci.IsCollapsed = (collapsed);
        }

        /// <summary>
        /// Collapses the col info records.
        /// </summary>
        /// <param name="columnIdx">The column index.</param>
        public void CollapseColInfoRecords(int columnIdx)
        {
            if (columnIdx == 0)
                return;
            ColumnInfoRecord previousCol = (ColumnInfoRecord)records[columnIdx - 1];
            ColumnInfoRecord currentCol = (ColumnInfoRecord)records[columnIdx];
            bool adjacentColumns = previousCol.LastColumn == currentCol.FirstColumn - 1;
            if (!adjacentColumns)
                return;

            bool columnsMatch =
                    previousCol.XFIndex == currentCol.XFIndex &&
                    previousCol.Options == currentCol.Options &&
                    previousCol.ColumnWidth == currentCol.ColumnWidth;

            if (columnsMatch)
            {
                previousCol.LastColumn = currentCol.LastColumn;
                records.Remove(columnIdx);
            }
        }

        /// <summary>
        /// Creates an outline Group for the specified columns.
        /// </summary>
        /// <param name="fromColumnIx">Group from this column (inclusive)</param>
        /// <param name="toColumnIx">Group to this column (inclusive)</param>
        /// <param name="indent">if true the Group will be indented by one level;if false indenting will be Removed by one level.</param>
        public void GroupColumnRange(int fromColumnIx, int toColumnIx, bool indent)
        {

            int colInfoSearchStartIdx = 0; // optimization to speed up the search for col infos
            for (int i = fromColumnIx; i <= toColumnIx; i++)
            {
                int level = 1;
                int colInfoIdx = FindColInfoIdx(i, colInfoSearchStartIdx);
                if (colInfoIdx != -1)
                {
                    level = GetColInfo(colInfoIdx).OutlineLevel;
                    if (indent)
                    {
                        level++;
                    }
                    else
                    {
                        level--;
                    }
                    level = Math.Max(0, level);
                    level = Math.Min(7, level);
                    colInfoSearchStartIdx = Math.Max(0, colInfoIdx - 1); // -1 just in case this column is collapsed later.
                }
                SetColumn(i, null, null, level, null, null);
            }

        }
        /// <summary>
        /// Finds the ColumnInfoRecord
        ///  which contains the specified columnIndex
        /// </summary>
        /// <param name="columnIndex">index of the column (not the index of the ColumnInfoRecord)</param>
        /// <returns>        /// <c>null</c>
        ///  if no column info found for the specified column
        ///  </returns>
        public ColumnInfoRecord FindColumnInfo(int columnIndex)
        {
            int nInfos = records.Count;
            for (int i = 0; i < nInfos; i++)
            {
                ColumnInfoRecord ci = GetColInfo(i);
                if (ci.ContainsColumn(columnIndex))
                {
                    return ci;
                }
            }
            return null;
        }
        private int FindColInfoIdx(int columnIx, int fromColInfoIdx)
        {
            if (columnIx < 0)
            {
                throw new ArgumentException("column parameter out of range: " + columnIx);
            }
            if (fromColInfoIdx < 0)
            {
                throw new ArgumentException("fromIdx parameter out of range: " + fromColInfoIdx);
            }

            for (int k = fromColInfoIdx; k < records.Count; k++)
            {
                ColumnInfoRecord ci = GetColInfo(k);
                if (ci.ContainsColumn(columnIx))
                {
                    return k;
                }
                if (ci.FirstColumn > columnIx)
                {
                    break;
                }
            }
            return -1;
        }
        /// <summary>
        /// Gets the max outline level.
        /// </summary>
        /// <value>The max outline level.</value>
        public int MaxOutlineLevel
        {
            get
            {
                int result = 0;
                int count = records.Count;
                for (int i = 0; i < count; i++)
                {
                    ColumnInfoRecord columnInfoRecord = GetColInfo(i);
                    result = Math.Max(columnInfoRecord.OutlineLevel, result);
                }
                return result;
            }
        }

        public int GetOutlineLevel(int columnIndex)
        {
            ColumnInfoRecord ci = FindColumnInfo(columnIndex);
            if (ci != null)
            {
                return ci.OutlineLevel;
            }
            else
            {
                return 0;
            }
        }
    }
}
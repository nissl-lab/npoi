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


using NPOI.Util;

namespace NPOI.HSSF.Record.Aggregates
{
    using System;
    using System.Text;
    using System.Collections.Generic;
    using NPOI.HSSF.Record;
    using NPOI.SS.Util;

    /// <summary>
    /// Manages various auxiliary records while constructing a RowRecordsAggregate
    /// @author Josh Micich
    /// </summary>
    [Serializable]
    public class SharedValueManager
    {

        private class SharedFormulaGroup
        {
            private SharedFormulaRecord _sfr;
            private FormulaRecordAggregate[] _frAggs;
            private int _numberOfFormulas;
            /**
             * Coordinates of the first cell having a formula that uses this shared formula.
             * This is often <i>but not always</i> the top left cell in the range covered by
             * {@link #_sfr}
             */
            private CellReference _firstCell;
            internal CellReference FirstCell
            {
                get { return _firstCell; }
            }
            public SharedFormulaGroup(SharedFormulaRecord sfr, CellReference firstCell)
            {
                if (!sfr.IsInRange(firstCell.Row, firstCell.Col))
                {
                    throw new ArgumentException("First formula cell " + firstCell.FormatAsString()
                            + " is not shared formula range " + sfr.Range.ToString() + ".");
                }
                _sfr = sfr;
                _firstCell = firstCell;
                int width = sfr.LastColumn - sfr.FirstColumn + 1;
                int height = sfr.LastRow - sfr.FirstRow + 1;
                _frAggs = new FormulaRecordAggregate[width * height];
                _numberOfFormulas = 0;
            }

            public void Add(FormulaRecordAggregate agg)
            {
                if (_numberOfFormulas == 0)
                {
                    if (_firstCell.Row != agg.Row || _firstCell.Col != agg.Column)
                    {
                        throw new InvalidOperationException("shared formula coding error");
                    }
                }
                if (_numberOfFormulas >= _frAggs.Length)
                {
                    throw new Exception("Too many formula records for shared formula group");
                }
                _frAggs[_numberOfFormulas++] = agg;
            }

            public void UnlinkSharedFormulas()
            {
                for (int i = 0; i < _numberOfFormulas; i++)
                {
                    _frAggs[i].UnlinkSharedFormula();
                }
            }

            public SharedFormulaRecord SFR
            {
                get
                {
                    return _sfr;
                }
            }

            public override String ToString()
            {
                StringBuilder sb = new StringBuilder(64);
                sb.Append(GetType().Name).Append(" [");
                sb.Append(_sfr.Range.ToString());
                sb.Append("]");
                return sb.ToString();
            }

            /**
             * Note - the 'first cell' of a shared formula group is not always the top-left cell
             * of the enclosing range.
             * @return <c>true</c> if the specified coordinates correspond to the 'first cell'
             * of this shared formula group.
             */
            public bool IsFirstCell(int row, int column)
            {
                return _firstCell.Row == row && _firstCell.Col == column;
            }
        }

        public static readonly SharedValueManager EMPTY = new SharedValueManager(
                new SharedFormulaRecord[0], new CellReference[0], new List<ArrayRecord>(), new List<TableRecord>());
        private List<ArrayRecord> _arrayRecords;
        private List<TableRecord> _tableRecords;
        private Dictionary<SharedFormulaRecord, SharedFormulaGroup> _groupsBySharedFormulaRecord;
        /** cached for optimization purposes */
        [NonSerialized]
        private Dictionary<int, SharedFormulaGroup> _groupsCache;

        private SharedValueManager(SharedFormulaRecord[] sharedFormulaRecords,
                CellReference[] firstCells, List<ArrayRecord> arrayRecords, List<TableRecord> tableRecords)
        {
            int nShF = sharedFormulaRecords.Length;
            if (nShF != firstCells.Length)
            {
                throw new ArgumentException("array sizes don't match: " + nShF + "!=" + firstCells.Length + ".");
            }
            _arrayRecords = new List<ArrayRecord>();
            _arrayRecords.AddRange(arrayRecords);
            _tableRecords = tableRecords;
            Dictionary<SharedFormulaRecord, SharedFormulaGroup> m = new Dictionary<SharedFormulaRecord, SharedFormulaGroup>(nShF * 3 / 2);
            for (int i = 0; i < nShF; i++)
            {
                SharedFormulaRecord sfr = sharedFormulaRecords[i];
                m[sfr] = new SharedFormulaGroup(sfr, firstCells[i]);
            }
            _groupsBySharedFormulaRecord = m;
        }
        public static SharedValueManager CreateEmpty()
        {
            // Note - must create distinct instances because they are assumed to be mutable.
            return new SharedValueManager(
                new SharedFormulaRecord[0], new CellReference[0], new List<ArrayRecord>(), new List<TableRecord>());
        }
        /**
         * @param firstCells
         * @param recs list of sheet records (possibly Contains records for other parts of the Excel file)
         * @param startIx index of first row/cell record for current sheet
         * @param endIx one past index of last row/cell record for current sheet.  It is important
         * that this code does not inadvertently collect <c>SharedFormulaRecord</c>s from any other
         * sheet (which could happen if endIx is chosen poorly).  (see bug 44449)
         */
        public static SharedValueManager Create(SharedFormulaRecord[] sharedFormulaRecords,
                CellReference[] firstCells, List<ArrayRecord> arrayRecords, List<TableRecord> tableRecords)
        {
            if (sharedFormulaRecords.Length + firstCells.Length + arrayRecords.Count + tableRecords.Count < 1)
            {
                return EMPTY;
            }
            return new SharedValueManager(sharedFormulaRecords, firstCells, arrayRecords, tableRecords);
        }


        /**
         * @param firstCell as extracted from the {@link ExpPtg} from the cell's formula.
         * @return never <code>null</code>
         */
        public SharedFormulaRecord LinkSharedFormulaRecord(CellReference firstCell, FormulaRecordAggregate agg)
        {

            SharedFormulaGroup result = FindFormulaGroupForCell(firstCell);
            if (null == result)
            {
                throw new RuntimeException("Failed to find a matching shared formula record");
            }
            result.Add(agg);
            return result.SFR;
        }

        private SharedFormulaGroup FindFormulaGroupForCell(CellReference cellRef)
        {
            if (null == _groupsCache)
            {
                _groupsCache = new Dictionary<int, SharedFormulaGroup>(_groupsBySharedFormulaRecord.Count);
                foreach (SharedFormulaGroup group in _groupsBySharedFormulaRecord.Values)
                {
                    _groupsCache.Add(GetKeyForCache(group.FirstCell), group);
                }
            }
            int key=GetKeyForCache(cellRef);
            SharedFormulaGroup sfg = null;
            if (_groupsCache.ContainsKey(key))
            {
                sfg = _groupsCache[key];
            }
            
            return sfg;
        }

        private int GetKeyForCache(CellReference cellRef)
        {
            // The HSSF has a max of 2^16 rows and 2^8 cols
            return ((cellRef.Col + 1) << 16 | cellRef.Row);
        }

        //private static SharedFormulaGroup FindFormulaGroup(SharedFormulaGroup[] groups, CellReference firstCell)
        //{
        //    int row = firstCell.Row;
        //    int column = firstCell.Col;
        //    // Traverse the list of shared formulas and try to find the correct one for us

        //    // perhaps this could be optimised to some kind of binary search
        //    for (int i = 0; i < groups.Length; i++)
        //    {
        //        SharedFormulaGroup svg = groups[i];
        //        if (svg.IsFirstCell(row, column))
        //        {
        //            return svg;
        //        }
        //    }
        //    // TODO - fix file "15228.xls" so it opens in Excel after rewriting with POI
        //    throw new Exception("Failed to find a matching shared formula record");
        //}

        //private SharedFormulaGroup[] GetGroups()
        //{
        //    if (_groupsCache == null)
        //    {
        //        SharedFormulaGroup[] groups = new SharedFormulaGroup[_groupsBySharedFormulaRecord.Count];
        //        _groupsBySharedFormulaRecord.Values.CopyTo(groups, 0);
        //        Array.Sort(groups, SVGComparator); // make search behaviour more deterministic
        //        _groupsCache = groups;
        //    }
        //    return _groupsCache;
        //}

        [NonSerialized]
        private SharedFormulaGroupComparator SVGComparator = new SharedFormulaGroupComparator();
        private class SharedFormulaGroupComparator : Comparer<SharedFormulaGroup>
        {
            public override int Compare(SharedFormulaGroup a, SharedFormulaGroup b)
            {
                CellRangeAddress8Bit rangeA = a.SFR.Range;
                CellRangeAddress8Bit rangeB = b.SFR.Range;

                int cmp;
                cmp = rangeA.FirstRow - rangeB.FirstRow;
                if (cmp != 0)
                {
                    return cmp;
                }
                cmp = rangeA.FirstColumn - rangeB.FirstColumn;
                if (cmp != 0)
                {
                    return cmp;
                }
                return 0;
            }
        }

        /**
         * Gets the {@link SharedValueRecordBase} record if it should be encoded immediately after the
         * formula record Contained in the specified {@link FormulaRecordAggregate} agg.  Note - the
         * shared value record always appears after the first formula record in the group.  For arrays
         * and tables the first formula is always the in the top left cell.  However, since shared
         * formula groups can be sparse and/or overlap, the first formula may not actually be in the
         * top left cell.
         *
         * @return the SHRFMLA, TABLE or ARRAY record for the formula cell, if it is the first cell of
         * a table or array region. <code>null</code> if the formula cell is not shared/array/table,
         * or if the specified formula is not the the first in the group.
         */
        public SharedValueRecordBase GetRecordForFirstCell(FormulaRecordAggregate agg)
        {
            CellReference firstCell = agg.FormulaRecord.Formula.ExpReference;
            // perhaps this could be optimised by consulting the (somewhat unreliable) isShared flag
            // and/or distinguishing between tExp and tTbl.
            if (firstCell == null)
            {
                // not a shared/array/table formula
                return null;
            }


            int row = firstCell.Row;
            int column = firstCell.Col;
            if (agg.Row != row || agg.Column != column)
            {
                // not the first formula cell in the group
                return null;
            }
            //SharedFormulaGroup[] groups = GetGroups();
            //for (int i = 0; i < groups.Length; i++)
            //{
            //    // note - logic for Finding correct shared formula group is slightly
            //    // more complicated since the first cell
            //    SharedFormulaGroup sfg = groups[i];
            //    if (sfg.IsFirstCell(row, column))
            //    {
            //        return sfg.SFR;
            //    }
            //}
            if (!(_groupsBySharedFormulaRecord.Count==0))
            {
                SharedFormulaGroup sfg = FindFormulaGroupForCell(firstCell);
                if (null != sfg)
                {
                    return sfg.SFR;
                }
            }
            // Since arrays and tables cannot be sparse (all cells in range participate)
            // The first cell will be the top left in the range.  So we can match the
            // ARRAY/TABLE record directly.

            for (int i = 0; i < _tableRecords.Count; i++)
            {
                TableRecord tr = _tableRecords[i];
                if (tr.IsFirstCell(row, column))
                {
                    return tr;
                }
            }
            foreach (ArrayRecord ar in _arrayRecords)
            {
                if (ar.IsFirstCell(row, column))
                {
                    return ar;
                }
            }
            return null;
        }

        /**
         * Converts all {@link FormulaRecord}s handled by <c>sharedFormulaRecord</c>
         * to plain unshared formulas
         */
        public void Unlink(SharedFormulaRecord sharedFormulaRecord)
        {
            SharedFormulaGroup svg = _groupsBySharedFormulaRecord[sharedFormulaRecord];
            _groupsBySharedFormulaRecord.Remove(sharedFormulaRecord);
            _groupsCache = null; // be sure to reset cached value
            if (svg == null)
            {
                throw new InvalidOperationException("Failed to find formulas for shared formula");
            }
            svg.UnlinkSharedFormulas();
        }

        /**
 * Add specified Array Record.
 */
        public void AddArrayRecord(ArrayRecord ar)
        {
            // could do a check here to make sure none of the ranges overlap
            _arrayRecords.Add(ar);
        }

        /**
         * Removes the {@link ArrayRecord} for the cell group containing the specified cell.
         * The caller should clear (set blank) all cells in the returned range.
         * @return the range of the array formula which was just removed. Never <code>null</code>.
         */
        public CellRangeAddress8Bit RemoveArrayFormula(int rowIndex, int columnIndex)
        {
            foreach (ArrayRecord ar in _arrayRecords)
            {
                if (ar.IsInRange(rowIndex, columnIndex))
                {
                    _arrayRecords.Remove(ar);
                    return ar.Range;
                }
            }
            String ref1 = new CellReference(rowIndex, columnIndex, false, false).FormatAsString();
            throw new ArgumentException("Specified cell " + ref1
                    + " is not part of an array formula.");
        }

        /**
         * @return the shared ArrayRecord identified by (firstRow, firstColumn). never <code>null</code>.
         */
        public ArrayRecord GetArrayRecord(int firstRow, int firstColumn)
        {
            foreach (ArrayRecord ar in _arrayRecords)
            {
                if (ar.IsFirstCell(firstRow, firstColumn))
                {
                    return ar;
                }
            }
            return null;
        }

    }
}
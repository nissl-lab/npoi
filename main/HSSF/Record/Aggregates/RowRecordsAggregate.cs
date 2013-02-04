
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.Record.Aggregates
{
    using System;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.SS.Formula;
    using NPOI.HSSF.Model;
    using System.Collections.Generic;
    using NPOI.HSSF.Record.Chart;

    /**
     *
     * @author  andy
     * @author Jason Height (jheight at chariot dot net dot au)
     */

    public class RowRecordsAggregate : RecordAggregate
    {
        private int firstrow = -1;
        private int lastrow = -1;
        private SortedList _rowRecords;
        //private int size = 0;
        private ValueRecordsAggregate _valuesAgg;
        private List<Record> _unknownRecords;
        private SharedValueManager _sharedValueManager;

        // Cache values to speed up performance of
        // getStartRowNumberForBlock / getEndRowNumberForBlock, see Bugzilla 47405
        private RowRecord[] _rowRecordValues = null;


        /** Creates a new instance of ValueRecordsAggregate */

        public RowRecordsAggregate()
            : this(SharedValueManager.CreateEmpty())
        {
        }

        public CellValueRecordInterface[] GetValueRecords()
        {
            return _valuesAgg.GetValueRecords();
        }
        private RowRecordsAggregate(SharedValueManager svm)
        {
            _rowRecords = new SortedList();
            _valuesAgg = new ValueRecordsAggregate();
            _unknownRecords = new List<Record>();
            _sharedValueManager = svm;
        }
        private int VisitRowRecordsForBlock(int blockIndex, RecordVisitor rv)
        {
            int startIndex = blockIndex * DBCellRecord.BLOCK_SIZE;
            int endIndex = startIndex + DBCellRecord.BLOCK_SIZE;

            IEnumerator rowIterator = _rowRecords.Values.GetEnumerator();

            //Given that we basically iterate through the rows in order,
            //For a performance improvement, it would be better to return an instance of
            //an iterator and use that instance throughout, rather than recreating one and
            //having to move it to the right position.
            int i = 0;
            for (; i < startIndex && rowIterator.MoveNext(); i++) ;

            int result = 0;
            while (rowIterator.MoveNext() && (i++ < endIndex))
            {
                Record rec = (Record)rowIterator.Current;
                result += rec.RecordSize;
                rv.VisitRecord(rec);
            }
            return result;
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {

            PositionTrackingVisitor stv = new PositionTrackingVisitor(rv, 0);
            //DBCells are serialized before row records.
            int blockCount = this.RowBlockCount;
            for (int blockIndex = 0; blockIndex < blockCount; blockIndex++)
            {
                // Serialize a block of rows.
                // Hold onto the position of the first row in the block
                int pos = 0;
                // Hold onto the size of this block that was serialized
                int rowBlockSize = VisitRowRecordsForBlock(blockIndex, rv);
                pos += rowBlockSize;
                // Serialize a block of cells for those rows
                int startRowNumber = GetStartRowNumberForBlock(blockIndex);
                int endRowNumber = GetEndRowNumberForBlock(blockIndex);
                DBCellRecord cellRecord = new DBCellRecord();
                // Note: Cell references start from the second row...
                int cellRefOffset = (rowBlockSize - RowRecord.ENCODED_SIZE);
                for (int row = startRowNumber; row <= endRowNumber; row++)
                {
                    if (_valuesAgg.RowHasCells(row))
                    {
                        stv.Position = 0;
                        _valuesAgg.VisitCellsForRow(row, stv);
                        int rowCellSize = stv.Position;
                        pos += rowCellSize;
                        // Add the offset to the first cell for the row into the
                        // DBCellRecord.
                        cellRecord.AddCellOffset((short)cellRefOffset);
                        cellRefOffset = rowCellSize;
                    }
                }
                // Calculate Offset from the start of a DBCellRecord to the first Row
                cellRecord.RowOffset = (pos);
                rv.VisitRecord(cellRecord);
            }
            for (int i = 0; i < _unknownRecords.Count; i++)
            {
                // Potentially breaking the file here since we don't know exactly where to write these records
                rv.VisitRecord((Record)_unknownRecords[i]);
            }
        }
        /**
         * @param rs record stream with all {@link SharedFormulaRecord}
         * {@link ArrayRecord}, {@link TableRecord} {@link MergeCellsRecord} Records removed
         */
        public RowRecordsAggregate(RecordStream rs, SharedValueManager svm)
            : this(svm)
        {
            while (rs.HasNext())
            {
                Record rec = rs.GetNext();
                switch (rec.Sid)
                {
                    case RowRecord.sid:
                        InsertRow((RowRecord)rec);
                        continue;
                    case DConRefRecord.sid:
                        AddUnknownRecord(rec);
                        continue;
                    case DBCellRecord.sid:
                        // end of 'Row Block'.  Should only occur after cell records
                        // ignore DBCELL records because POI generates them upon re-serialization
                        continue;
                }
                if (rec is UnknownRecord)
                {
                    // might need to keep track of where exactly these belong
                    AddUnknownRecord((UnknownRecord)rec);

                    while (rs.PeekNextSid() == ContinueRecord.sid)
                    {
                        AddUnknownRecord(rs.GetNext());
                    }
                    continue;
                }
       			if (rec is MulBlankRecord) {
    			    _valuesAgg.AddMultipleBlanks((MulBlankRecord) rec);
				    continue;
			    }

                if (!(rec is CellValueRecordInterface))
                {
                    //TODO: correct it, SeriesIndexRecord will appear in a separate chart sheet that contains a single chart
                    // rule SERIESDATA = Dimensions 3(SIIndex *(Number / BoolErr / Blank / Label))
                    if (rec.Sid == SeriesIndexRecord.sid)
                    {
                        AddUnknownRecord(rec);
                        continue;
                    }
                    throw new InvalidOperationException("Unexpected record type (" + rec.GetType().Name + ")");

                }
                _valuesAgg.Construct((CellValueRecordInterface)rec, rs, svm);
            }
        }
        /**
  * Handles UnknownRecords which appear within the row/cell records
  */
        private void AddUnknownRecord(Record rec)
        {
            // ony a few distinct record IDs are encountered by the existing POI test cases:
            // 0x1065 // many
            // 0x01C2 // several
            // 0x0034 // few
            // No documentation could be found for these

            // keep the unknown records for re-serialization
            _unknownRecords.Add(rec);
        }
        public void InsertRow(RowRecord row)
        {
            _rowRecords[row.RowNumber] = row;
            // Clear the cached values
            _rowRecordValues = null; 


            if (row.RowNumber < firstrow|| firstrow == -1)
            {
                firstrow = row.RowNumber;
            }
            if (row.RowNumber > lastrow|| lastrow == -1)
            {
                lastrow = row.RowNumber;
            }
        }

        public void RemoveRow(RowRecord row)
        {
            int rowIndex = row.RowNumber;
            _valuesAgg.RemoveAllCellsValuesForRow(rowIndex);
            int key = rowIndex;
            RowRecord rr = (RowRecord)_rowRecords[key];
            _rowRecords.Remove(key);
            if (rr == null)
            {
                throw new Exception("Invalid row index (" + key + ")");
            }
            if (row != rr)
            {
                _rowRecords[key] = rr;
                throw new Exception("Attempt to remove row that does not belong to this sheet");
            }
            // Clear the cached values
            _rowRecordValues = null;
        }
        public void InsertCell(CellValueRecordInterface cvRec)
        {
            _valuesAgg.InsertCell(cvRec);
        }
        public void RemoveCell(CellValueRecordInterface cvRec)
        {
            if (cvRec is FormulaRecordAggregate)
            {
                ((FormulaRecordAggregate)cvRec).NotifyFormulaChanging();
            }
            _valuesAgg.RemoveCell(cvRec);
        }
        public RowRecord GetRow(int rowIndex)
        {
            // Row must be between 0 and 65535
            if (rowIndex < 0 || rowIndex > 65535)
            {
                throw new ArgumentException("The row number must be between 0 and 65535");
            }
            return (RowRecord)_rowRecords[rowIndex];
        }
        public FormulaRecordAggregate CreateFormula(int row, int col)
        {
            FormulaRecord fr = new FormulaRecord();
            fr.Row=(row);
            fr.Column=((short)col);
            return new FormulaRecordAggregate(fr, null, _sharedValueManager);
        }
        public int PhysicalNumberOfRows
        {
            get
            {
                return _rowRecords.Count;
            }
        }

        public int FirstRowNum
        {
            get
            {
                return firstrow;
            }
        }

        public int LastRowNum
        {
            get
            {
                return lastrow;
            }
        }

        /** Returns the number of row blocks.
         * <p/>The row blocks are goupings of rows that contain the DBCell record
         * after them
         */
        public int RowBlockCount
        {
            get
            {
                int size = _rowRecords.Count / DBCellRecord.BLOCK_SIZE;
                if ((_rowRecords.Count % DBCellRecord.BLOCK_SIZE) != 0)
                    size++;
                return size;
            }
        }

        public int GetRowBlockSize(int block)
        {
            return 20 * GetRowCountForBlock(block);
        }

        /** Returns the number of physical rows within a block*/
        public int GetRowCountForBlock(int block)
        {
            int startIndex = block * DBCellRecord.BLOCK_SIZE;
            int endIndex = startIndex + DBCellRecord.BLOCK_SIZE - 1;
            if (endIndex >= _rowRecords.Count)
                endIndex = _rowRecords.Count - 1;

            return endIndex - startIndex + 1;
        }

        /** Returns the physical row number of the first row in a block*/
        public int GetStartRowNumberForBlock(int block)
        {
            //Given that we basically iterate through the rows in order,
            //For a performance improvement, it would be better to return an instance of
            //an iterator and use that instance throughout, rather than recreating one and
            //having to move it to the right position.

            
            int startIndex = block * DBCellRecord.BLOCK_SIZE;

            if (_rowRecordValues == null)
            {
                _rowRecordValues=new RowRecord[_rowRecords.Count];
                _rowRecords.Values.CopyTo(_rowRecordValues,0);
            }
            try
            {
                return _rowRecordValues[startIndex].RowNumber;
            }
            catch (IndexOutOfRangeException)
            {
                throw new Exception("Did not find start row for block " + block);
            }
        }

        /** Returns the physical row number of the end row in a block*/
        public int GetEndRowNumberForBlock(int block)
        {
            int endIndex = ((block + 1) * DBCellRecord.BLOCK_SIZE) - 1;
            if (endIndex >= _rowRecords.Count)
                endIndex = _rowRecords.Count - 1;

            if (_rowRecordValues == null)
            {
                _rowRecordValues = new RowRecord[_rowRecords.Count];
                _rowRecords.Values.CopyTo(_rowRecordValues,0);
            }

            try
            {
                return _rowRecordValues[endIndex].RowNumber;
            }
            catch (IndexOutOfRangeException)
            {
                throw new Exception("Did not find end row for block " + block);
            }
        }

        public IEnumerator GetEnumerator()
        {
            return _rowRecords.Values.GetEnumerator();
        }


        public int FindStartOfRowOutlineGroup(int row)
        {
            // Find the start of the Group.
            RowRecord rowRecord = this.GetRow(row);
            int level = rowRecord.OutlineLevel;
            int currentRow = row;
            while (this.GetRow(currentRow) != null)
            {
                rowRecord = this.GetRow(currentRow);
                if (rowRecord.OutlineLevel < level)
                    return currentRow + 1;
                currentRow--;
            }

            return currentRow + 1;
        }

        public int FindEndOfRowOutlineGroup(int row)
        {
            int level = GetRow(row).OutlineLevel;
            int currentRow;
            for (currentRow = row; currentRow < this.LastRowNum; currentRow++)
            {
                if (GetRow(currentRow) == null || GetRow(currentRow).OutlineLevel < level)
                {
                    break;
                }
            }

            return currentRow - 1;
        }

        public int WriteHidden(RowRecord rowRecord, int row, bool hidden)
        {
            int level = rowRecord.OutlineLevel;
            while (rowRecord != null && this.GetRow(row).OutlineLevel >= level)
            {
                rowRecord.ZeroHeight = (hidden);
                row++;
                rowRecord = this.GetRow(row);
            }
            return row - 1;
        }

        public void CollapseRow(int rowNumber)
        {

            // Find the start of the Group.
            int startRow = FindStartOfRowOutlineGroup(rowNumber);
            RowRecord rowRecord = GetRow(startRow);

            // Hide all the columns Until the end of the Group
            int lastRow = WriteHidden(rowRecord, startRow, true);

            // Write collapse field
            if (GetRow(lastRow + 1) != null)
            {
                GetRow(lastRow + 1).Colapsed = (true);
            }
            else
            {
                RowRecord row = CreateRow(lastRow + 1);
                row.Colapsed = (true);
                InsertRow(row);
            }
        }
        public DimensionsRecord CreateDimensions()
        {
            DimensionsRecord result = new DimensionsRecord();
            result.FirstRow=(firstrow);
            result.LastRow=(lastrow);
            result.FirstCol =_valuesAgg.FirstCellNum;
            result.LastCol = _valuesAgg.LastCellNum;
            return result;
        }

        /**
         * Create a row record.
         *
         * @param row number
         * @return RowRecord Created for the passed in row number
         * @see org.apache.poi.hssf.record.RowRecord
         */
        public static RowRecord CreateRow(int rowNumber)
        {
            return new RowRecord(rowNumber);
        }

        public IndexRecord CreateIndexRecord(int indexRecordOffset, int sizeOfInitialSheetRecords, int offsetDefaultColWidth)
        {
            IndexRecord result = new IndexRecord();
            result.FirstRow= firstrow;
            result.LastRowAdd1= lastrow+1;
            // Calculate the size of the records from the end of the BOF
            // and up to the RowRecordsAggregate...

            // Add the references to the DBCells in the IndexRecord (one for each block)
            // Note: The offsets are relative to the Workbook BOF. Assume that this is
            // 0 for now.....

            int blockCount = RowBlockCount;
            // Calculate the size of this IndexRecord
            int indexRecSize = IndexRecord.GetRecordSizeForBlockCount(blockCount);

            int currentOffset = indexRecordOffset + indexRecSize + sizeOfInitialSheetRecords;

            for (int block = 0; block < blockCount; block++)
            {
                // each row-block has a DBCELL record.
                // The offset of each DBCELL record needs to be updated in the INDEX record

                // account for row records in this row-block
                currentOffset += GetRowBlockSize(block);
                // account for cell value records after those
                currentOffset += _valuesAgg.GetRowCellBlockSize(
                        GetStartRowNumberForBlock(block), GetEndRowNumberForBlock(block));

                // currentOffset is now the location of the DBCELL record for this row-block
                result.AddDbcell(currentOffset);
                // Add space required to write the DBCELL record (whose reference was just added).
                currentOffset += (8 + (GetRowCountForBlock(block) * 2));
            }
            return result;
        }

        public bool IsRowGroupCollapsed(int row)
        {
            int collapseRow = FindEndOfRowOutlineGroup(row) + 1;

            if (GetRow(collapseRow) == null)
                return false;
            else
                return GetRow(collapseRow).Colapsed;
        }

        public void ExpandRow(int rowNumber)
        {
            int idx = rowNumber;
            if (idx == -1)
                return;

            // If it is already expanded do nothing.
            if (!IsRowGroupCollapsed(idx))
                return;

            // Find the start of the Group.
            int startIdx = FindStartOfRowOutlineGroup(idx);
            RowRecord row = GetRow(startIdx);

            // Find the end of the Group.
            int endIdx = FindEndOfRowOutlineGroup(idx);

            // expand:
            // collapsed bit must be UnSet
            // hidden bit Gets UnSet _if_ surrounding Groups are expanded you can determine
            //   this by looking at the hidden bit of the enclosing Group.  You will have
            //   to look at the start and the end of the current Group to determine which
            //   is the enclosing Group
            // hidden bit only is altered for this outline level.  ie.  don't Un-collapse contained Groups
            if (!IsRowGroupHiddenByParent(idx))
            {
                for (int i = startIdx; i <= endIdx; i++)
                {
                    if (row.OutlineLevel == GetRow(i).OutlineLevel)
                        GetRow(i).ZeroHeight = (false);
                    else if (!IsRowGroupCollapsed(i))
                        GetRow(i).ZeroHeight = (false);
                }
            }

            // Write collapse field
            GetRow(endIdx + 1).Colapsed = (false);
        }

        public void UpdateFormulasAfterRowShift(FormulaShifter formulaShifter, int currentExternSheetIndex)
        {
            _valuesAgg.UpdateFormulasAfterRowShift(formulaShifter, currentExternSheetIndex);
        }

        public bool IsRowGroupHiddenByParent(int row)
        {
            // Look out outline details of end
            int endLevel;
            bool endHidden;
            int endOfOutlineGroupIdx = FindEndOfRowOutlineGroup(row);
            if (GetRow(endOfOutlineGroupIdx + 1) == null)
            {
                endLevel = 0;
                endHidden = false;
            }
            else
            {
                endLevel = GetRow(endOfOutlineGroupIdx + 1).OutlineLevel;
                endHidden = GetRow(endOfOutlineGroupIdx + 1).ZeroHeight;
            }

            // Look out outline details of start
            int startLevel;
            bool startHidden;
            int startOfOutlineGroupIdx = FindStartOfRowOutlineGroup(row);
            if (startOfOutlineGroupIdx - 1 < 0 || GetRow(startOfOutlineGroupIdx - 1) == null)
            {
                startLevel = 0;
                startHidden = false;
            }
            else
            {
                startLevel = GetRow(startOfOutlineGroupIdx - 1).OutlineLevel;
                startHidden = GetRow(startOfOutlineGroupIdx - 1).ZeroHeight;
            }

            if (endLevel > startLevel)
            {
                return endHidden;
            }
            else
            {
                return startHidden;
            }
        }

    }

}
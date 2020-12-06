
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
    using NPOI.SS.Formula.PTG;
    using System.Collections.Generic;

    /**
     *
     * Aggregate value records toGether.  Things are easier to handle that way.
     *
     * @author  andy
     * @author  Glen Stampoultzis (glens at apache.org)
     * @author Jason Height (jheight at chariot dot net dot au)
     */

    public class ValueRecordsAggregate : IEnumerable<CellValueRecordInterface>
    {
        const int MAX_ROW_INDEX = 0XFFFF;
        const int INDEX_NOT_SET = -1;
        public const short sid = -1001; // 1000 clashes with RowRecordsAggregate
        int firstcell = INDEX_NOT_SET;
        int lastcell = INDEX_NOT_SET;
        CellValueRecordInterface[][] records;

        /** Creates a new instance of ValueRecordsAggregate */

        public ValueRecordsAggregate() :
            this(INDEX_NOT_SET, INDEX_NOT_SET, new CellValueRecordInterface[30][]) // We start with 30 Rows.
        {

        }

        private ValueRecordsAggregate(int firstCellIx, int lastCellIx, CellValueRecordInterface[][] pRecords)
        {
            firstcell = firstCellIx;
            lastcell = lastCellIx;
            records = pRecords;
        }

        public void InsertCell(CellValueRecordInterface cell)
        {
            int column = cell.Column;
            int row = cell.Row;
            if (row >= records.Length)
            {
                CellValueRecordInterface[][] oldRecords = records;
                int newSize = oldRecords.Length * 2;
                if (newSize < row + 1) newSize = row + 1;
                records = new CellValueRecordInterface[newSize][];
                Array.Copy(oldRecords, 0, records, 0, oldRecords.Length);
            }

            object objRowCells = records[row];
            if (objRowCells == null)
            {
                int newSize = column + 1;
                if (newSize < 10) newSize = 10;
                objRowCells = new CellValueRecordInterface[newSize];
                records[row] = (CellValueRecordInterface[]) objRowCells;
            }

            CellValueRecordInterface[] rowCells = (CellValueRecordInterface[]) objRowCells;
            if (column >= rowCells.Length)
            {
                CellValueRecordInterface[] oldRowCells = rowCells;
                int newSize = oldRowCells.Length * 2;
                if (newSize < column + 1) newSize = column + 1;
                // if(newSize>257) newSize=257; // activate?
                rowCells = new CellValueRecordInterface[newSize];
                Array.Copy(oldRowCells, 0, rowCells, 0, oldRowCells.Length);
                records[row] = rowCells;
            }

            rowCells[column] = cell;

            if ((column < firstcell) || (firstcell == -1))
            {
                firstcell = column;
            }

            if ((column > lastcell) || (lastcell == -1))
            {
                lastcell = column;
            }
        }

        public void RemoveCell(CellValueRecordInterface cell)
        {
            if (cell == null)
            {
                throw new ArgumentException("cell must not be null");
            }

            int row = cell.Row;
            if (row >= records.Length)
            {
                throw new Exception("cell row is out of range");
            }

            CellValueRecordInterface[] rowCells = records[row];
            if (rowCells == null)
            {
                throw new Exception("cell row is already empty");
            }

            int column = cell.Column;
            if (column >= rowCells.Length)
            {
                throw new Exception("cell column is out of range");
            }

            rowCells[column] = null;

        }

        public void RemoveAllCellsValuesForRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex > MAX_ROW_INDEX)
            {
                throw new ArgumentException("Specified rowIndex " + rowIndex
                                                                  + " is outside the allowable range (0.." +
                                                                  MAX_ROW_INDEX + ")");
            }

            if (rowIndex >= records.Length)
            {
                // this can happen when the client code has created a row,
                // and then removes/replaces it before adding any cells. (see bug 46312)
                return;
            }

            records[rowIndex] = null;
        }

        public int PhysicalNumberOfCells
        {
            get
            {
                int count = 0;
                for (int r = 0; r < records.Length; r++)
                {
                    CellValueRecordInterface[] rowCells = records[r];
                    if (rowCells != null)
                        for (short c = 0; c < rowCells.Length; c++)
                        {
                            if (rowCells[c] != null) count++;
                        }
                }

                return count;
            }
        }

        public int FirstCellNum
        {
            get { return firstcell; }
        }

        public int LastCellNum
        {
            get { return lastcell; }
        }

        public void AddMultipleBlanks(MulBlankRecord mbr)
        {
            for (int j = 0; j < mbr.NumColumns; j++)
            {
                BlankRecord br = new BlankRecord();

                br.Column = j + mbr.FirstColumn;
                br.Row = mbr.Row;
                br.XFIndex = (mbr.GetXFAt(j));
                InsertCell(br);
            }
        }

        private MulBlankRecord CreateMBR(CellValueRecordInterface[] cellValues, int startIx, int nBlank)
        {

            short[] xfs = new short[nBlank];
            for (int i = 0; i < xfs.Length; i++)
            {
                xfs[i] = ((BlankRecord) cellValues[startIx + i]).XFIndex;
            }

            int rowIx = cellValues[startIx].Row;
            return new MulBlankRecord(rowIx, startIx, xfs);
        }

        public void Construct(CellValueRecordInterface rec, RecordStream rs, SharedValueManager sfh)
        {
            if (rec is FormulaRecord)
            {
                FormulaRecord formulaRec = (FormulaRecord) rec;
                // read optional cached text value
                StringRecord cachedText = null;
                Type nextClass = rs.PeekNextClass();
                if (nextClass == typeof(StringRecord))
                {
                    cachedText = (StringRecord) rs.GetNext();
                }
                else
                {
                    cachedText = null;
                }

                InsertCell(new FormulaRecordAggregate(formulaRec, cachedText, sfh));
            }
            else
            {
                InsertCell(rec);
            }
        }

        /**
         * Sometimes the shared formula flag "seems" to be erroneously Set, in which case there is no 
         * call to <c>SharedFormulaRecord.ConvertSharedFormulaRecord</c> and hence the 
         * <c>ParsedExpression</c> field of this <c>FormulaRecord</c> will not Get updated.<br/>
         * As it turns out, this is not a problem, because in these circumstances, the existing value
         * for <c>ParsedExpression</c> is perfectly OK.<p/>
         * 
         * This method may also be used for Setting breakpoints to help diagnose Issues regarding the
         * abnormally-Set 'shared formula' flags. 
         * (see TestValueRecordsAggregate.testSpuriousSharedFormulaFlag()).<p/>
         * 
         * The method currently does nothing but do not delete it without Finding a nice home for this 
         * comment.
         */
        static void HandleMissingSharedFormulaRecord(FormulaRecord formula)
        {
            // could log an info message here since this is a fairly Unusual occurrence.
        }


        /** Tallies a count of the size of the cell records
         *  that are attached to the rows in the range specified.
         */
        public int GetRowCellBlockSize(int startRow, int endRow)
        {
            ValueEnumerator itr = new ValueEnumerator(ref records, startRow, endRow);
            int size = 0;
            while (itr.MoveNext())
            {
                CellValueRecordInterface cell = (CellValueRecordInterface) itr.Current;
                int row = cell.Row;
                if (row > endRow)
                    break;
                if ((row >= startRow) && (row <= endRow))
                    size += ((RecordBase) cell).RecordSize;
            }

            return size;
        }

        /** Returns true if the row has cells attached to it */
        public bool RowHasCells(int row)
        {
            if (row > records.Length - 1) //previously this said row > records.Length which means if 
                return false; // if records.Length == 60 and I pass "60" here I Get array out of bounds
            CellValueRecordInterface[] rowCells = records[row]; //because a 60 Length array has the last index = 59
            if (rowCells == null) return false;
            for (int col = 0; col < rowCells.Length; col++)
            {
                if (rowCells[col] != null) return true;
            }

            return false;
        }

        public void UpdateFormulasAfterRowShift(FormulaShifter shifter, int currentExternSheetIndex)
        {
            for (int i = 0; i < records.Length; i++)
            {
                CellValueRecordInterface[] rowCells = records[i];
                if (rowCells == null)
                {
                    continue;
                }

                for (int j = 0; j < rowCells.Length; j++)
                {
                    CellValueRecordInterface cell = rowCells[j];
                    if (cell is FormulaRecordAggregate)
                    {
                        FormulaRecordAggregate fra = (FormulaRecordAggregate) cell;
                        Ptg[] ptgs = fra.FormulaTokens; // needs clone() inside this getter?
                        Ptg[] ptgs2 = ((FormulaRecordAggregate) cell).FormulaRecord
                            .ParsedExpression; // needs clone() inside this getter?

                        if (shifter.AdjustFormula(ptgs, currentExternSheetIndex))
                        {
                            fra.SetParsedExpression(ptgs);
                        }
                    }
                }
            }
        }

        public void VisitCellsForRow(int rowIndex, RecordVisitor rv)
        {

            CellValueRecordInterface[] rowCells = records[rowIndex];
            if (rowCells == null)
            {
                throw new ArgumentException("Row [" + rowIndex + "] is empty");
            }

            for (int i = 0; i < rowCells.Length; i++)
            {
                RecordBase cvr = (RecordBase) rowCells[i];
                if (cvr == null)
                {
                    continue;
                }

                int nBlank = CountBlanks(rowCells, i);
                if (nBlank > 1)
                {
                    rv.VisitRecord(CreateMBR(rowCells, i, nBlank));
                    i += nBlank - 1;
                }
                else if (cvr is RecordAggregate)
                {
                    RecordAggregate agg = (RecordAggregate) cvr;
                    agg.VisitContainedRecords(rv);
                }
                else
                {
                    rv.VisitRecord((Record) cvr);
                }
            }
        }

        static int CountBlanks(CellValueRecordInterface[] rowCellValues, int startIx)
        {
            int i = startIx;
            while (i < rowCellValues.Length)
            {
                CellValueRecordInterface cvr = rowCellValues[i];
                if (!(cvr is BlankRecord))
                {
                    break;
                }

                i++;
            }

            return i - startIx;
        }

        /** Serializes the cells that are allocated to a certain row range*/
        public int SerializeCellRow(int row, int offset, byte[] data)
        {
            ValueEnumerator itr = new ValueEnumerator(ref records, row, row);
            int pos = offset;

            while (itr.MoveNext())
            {
                CellValueRecordInterface cell = (CellValueRecordInterface) itr.Current;
                if (cell.Row != row)
                    break;
                pos += ((RecordBase) cell).Serialize(pos, data);
            }

            return pos - offset;
        }

        public IEnumerator<CellValueRecordInterface> GetEnumerator()
        {
            return new ValueEnumerator(ref records);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private class ValueEnumerator : IEnumerator<CellValueRecordInterface>
        {
            short nextColumn = -1;
            int nextRow, lastRow;

            CellValueRecordInterface[][] records;

            public ValueEnumerator(ref CellValueRecordInterface[][] _records)
            {
                this.records = _records;
                this.nextRow = 0;
                this.lastRow = _records.Length - 1;
                //FindNext();
            }

            public ValueEnumerator(ref CellValueRecordInterface[][] _records, int firstRow, int lastRow)
            {
                this.records = _records;
                this.nextRow = firstRow;
                this.lastRow = lastRow;
                //FindNext();
            }

            public bool MoveNext()
            {

                FindNext();
                return nextRow <= lastRow;
                ;
            }

            public CellValueRecordInterface Current
            {
                get
                {
                    CellValueRecordInterface o = this.records[nextRow][nextColumn];

                    return o;
                }
            }

            object IEnumerator.Current
            {
                get { return this.Current; }
            }

            public void Remove()
            {
                throw new InvalidOperationException("gibt's noch nicht");
            }

            private void FindNext()
            {
                nextColumn++;
                for (; nextRow <= lastRow; nextRow++)
                {
                    //previously this threw array out of bounds...
                    CellValueRecordInterface[] rowCells = (nextRow < records.Length) ? records[nextRow] : null;
                    if (rowCells == null)
                    {
                        // This row is empty
                        nextColumn = 0;
                        continue;
                    }

                    for (; nextColumn < rowCells.Length; nextColumn++)
                    {
                        if (rowCells[nextColumn] != null) return;
                    }

                    nextColumn = 0;
                }
            }

            public void Reset()
            {
                nextColumn = -1;
                nextRow = 0;
            }

            public void Dispose()
            {
                this.records = null;
            }
        }
    }
}
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
namespace NPOI.HSSF.EventUserModel
{
    using NPOI.HSSF.Record;
    using NPOI.HSSF.EventUserModel;
    using NPOI.HSSF.EventUserModel.DummyRecord;

    /// <summary>
    /// A HSSFListener which tracks rows and columns, and will
    /// trigger your HSSFListener for all rows and cells,
    /// even the ones that aren't actually stored in the file.
    /// This allows your code to have a more "Excel" like
    /// view of the data in the file, and not have to worry
    /// (as much) about if a particular row/cell Is in the
    /// file, or was skipped from being written as it was
    /// blank.
    /// </summary>
    public class MissingRecordAwareHSSFListener : IHSSFListener
    {
        private IHSSFListener childListener;
        // Need to have different counters for cell rows and
        //  row rows, as you sometimes get a RowRecord in the
        //  middle of some cells, and that'd break everything
        private int lastRowRow;

        private int lastCellRow;
        private int lastCellColumn;

        /// <summary>
        /// Constructs a new MissingRecordAwareHSSFListener, which
        /// will fire ProcessRecord on the supplied child
        /// HSSFListener for all Records, and missing records.
        /// </summary>
        /// <param name="listener">The HSSFListener to pass records on to</param>
        public MissingRecordAwareHSSFListener(IHSSFListener listener)
        {
            ResetCounts();
            childListener = listener;
        }

        /// <summary>
        /// Process an HSSF Record. Called when a record occurs in an HSSF file.
        /// </summary>
        /// <param name="record"></param>
        public void ProcessRecord(Record record)
        {
            int thisRow;
            int thisColumn;
            CellValueRecordInterface[] expandedRecords = null;

            if (record is CellValueRecordInterface)
            {
                CellValueRecordInterface valueRec = (CellValueRecordInterface)record;
                thisRow = valueRec.Row;
                thisColumn = valueRec.Column;
            }
            else
            {
                if (record is StringRecord)
                {
                    //it contains only cashed result of the previous FormulaRecord evaluation
                    childListener.ProcessRecord(record);
                    return;
                }
                thisRow = -1;
                thisColumn = -1;

                switch (record.Sid)
                {
                    // the BOFRecord can represent either the beginning of a sheet or the workbook
                    case BOFRecord.sid:
                        BOFRecord bof = (BOFRecord)record;
                        if (bof.Type == BOFRecordType.Workbook || bof.Type == BOFRecordType.Worksheet)
                        {
                            // Reset the row and column counts - new workbook / worksheet
                            ResetCounts();
                        }
                        break;
                    case RowRecord.sid:
                        RowRecord rowrec = (RowRecord)record;
                        //Console.WriteLine("Row " + rowrec.RowNumber + " found, first column at "
                        //        + rowrec.GetFirstCol() + " last column at " + rowrec.GetLastCol());

                        // If there's a jump in rows, fire off missing row records
                        if (lastRowRow + 1 < rowrec.RowNumber)
                        {
                            for (int i = (lastRowRow + 1); i < rowrec.RowNumber; i++)
                            {
                                MissingRowDummyRecord dr = new MissingRowDummyRecord(i);
                                childListener.ProcessRecord(dr);
                            }
                        }

                        // Record this as the last row we saw
                        lastRowRow = rowrec.RowNumber;
                        break;
                    case SharedFormulaRecord.sid:
                        // SharedFormulaRecord occurs after the first FormulaRecord of the cell range.
                        // There are probably (but not always) more cell records after this
                        // - so don't fire off the LastCellOfRowDummyRecord yet
                        childListener.ProcessRecord(record);
                        return;
                    case MulBlankRecord.sid:
                        // These appear in the middle of the cell records, to
                        //  specify that the next bunch are empty but styled
                        // Expand this out into multiple blank cells
                        MulBlankRecord mbr = (MulBlankRecord)record;
                        expandedRecords = RecordFactory.ConvertBlankRecords(mbr);
                        break;
                    case MulRKRecord.sid:
                        // This is multiple consecutive number cells in one record
                        // Exand this out into multiple regular number cells
                        MulRKRecord mrk = (MulRKRecord)record;
                        expandedRecords = RecordFactory.ConvertRKRecords(mrk);
                        break;
                    case NoteRecord.sid:
                        NoteRecord nrec = (NoteRecord)record;
                        thisRow = nrec.Row;
                        thisColumn = nrec.Column;
                        break;
                    default:
                        //Console.WriteLine(record.GetClass());
                        break;
                }
            }

            // First part of expanded record handling
            if (expandedRecords != null && expandedRecords.Length > 0)
            {
                thisRow = expandedRecords[0].Row;
                thisColumn = expandedRecords[0].Column;
            }

            // If we're on cells, and this cell isn't in the same
            //  row as the last one, then fire the 
            //  dummy end-of-row records
            if (thisRow != lastCellRow && lastCellRow > -1)
            {
                for (int i = lastCellRow; i < thisRow; i++)
                {
                    int cols = -1;
                    if (i == lastCellRow)
                    {
                        cols = lastCellColumn;
                    }
                    childListener.ProcessRecord(new LastCellOfRowDummyRecord(i, cols));
                }
            }

            // If we've just finished with the cells, then fire the
            // final dummy end-of-row record
            if (lastCellRow != -1 && lastCellColumn != -1 && thisRow == -1)
            {
                childListener.ProcessRecord(new LastCellOfRowDummyRecord(lastCellRow, lastCellColumn));

                lastCellRow = -1;
                lastCellColumn = -1;
            }

            // If we've moved onto a new row, the ensure we re-set
            //  the column counter
            if (thisRow != lastCellRow)
            {
                lastCellColumn = -1;
            }

            // If there's a gap in the cells, then fire
            //  the dummy cell records
            if (lastCellColumn != thisColumn - 1)
            {
                for (int i = lastCellColumn + 1; i < thisColumn; i++)
                {
                    childListener.ProcessRecord(new MissingCellDummyRecord(thisRow, i));
                }
            }

            // Next part of expanded record handling
            if (expandedRecords != null && expandedRecords.Length > 0)
            {
                thisColumn = expandedRecords[expandedRecords.Length - 1].Column;
            }


            // Update cell and row counts as needed
            if (thisColumn != -1)
            {
                lastCellColumn = thisColumn;
                lastCellRow = thisRow;
            }

            // Pass along the record(s)
            if (expandedRecords != null && expandedRecords.Length > 0)
            {
                foreach (CellValueRecordInterface r in expandedRecords)
                {
                    childListener.ProcessRecord((Record)r);
                }
            }
            else
            {
                childListener.ProcessRecord(record);
            }
        }
        private void ResetCounts()
        {
            lastRowRow = -1;
            lastCellRow = -1;
            lastCellColumn = -1;
        }
    }
}
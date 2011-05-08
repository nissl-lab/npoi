/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Model
{
    using System;
    using System.Collections;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.EventModel;
    using NPOI.HSSF.Model;
    using NPOI.SS.Util;
    using TestCases.HSSF.UserModel;
    using System.Collections.Generic;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record.Formula;

    /**
     * Unit Test for the Sheet class.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestClass]
    public class TestSheet
    {

        private static Record[] GetSheetRecords(InternalSheet s, int offset)
        {
            RecordInspector.RecordCollector rc = new RecordInspector.RecordCollector();
            s.VisitContainedRecords(rc, offset);
            return rc.Records;
        }
        private static InternalSheet CreateSheet(ArrayList inRecs)
        {
            return InternalSheet.CreateSheet(new RecordStream(inRecs, 0));
        }

        [TestMethod]
        public void TestCreateSheet()
        {
            // Check we're Adding row and cell aggregates
            ArrayList records = new ArrayList();
            records.Add(new BOFRecord());
            records.Add(new DimensionsRecord());
            records.Add(CreateWindow2Record());
            records.Add(new EOFRecord());
            InternalSheet sheet = CreateSheet(records);
            Record[] outRecs = GetSheetRecords(sheet, 0);

            int pos = 0;
            Assert.IsTrue(outRecs[pos++] is BOFRecord);
            Assert.IsTrue(outRecs[pos++] is IndexRecord);
            Assert.IsTrue(outRecs[pos++] is DimensionsRecord);
            Assert.IsTrue(outRecs[pos++] is WindowTwoRecord);
            Assert.IsTrue(outRecs[pos++] is EOFRecord);
        }

       private class MergedCellListener:RecordVisitor 
       {

            private int _count;
            public MergedCellListener() {
                _count = 0;
            }
            public void VisitRecord(Record r) {
                if (r is MergeCellsRecord) {
                    _count++;
                }
            }
            public int Count
            {
                get
                {
                    return _count;
                }
            }
        }

        [TestMethod]
        public void TestAddMergedRegion()
        {
            InternalSheet sheet = InternalSheet.CreateSheet();
            int regionsToAdd = 4096;
            int startRecords = sheet.Records.Count;

            //simple Test that Adds a load of regions
            for (int n = 0; n < regionsToAdd; n++)
            {
                int index = sheet.AddMergedRegion(0, 0, 1, 1);
                Assert.AreEqual(index, n, "Merged region index expected to be " + n + " got " + index);
            }

            //test all the regions were indeed Added
            Assert.AreEqual(sheet.NumMergedRegions, regionsToAdd);

            //test that the regions were spread out over the appropriate number of records
            MergedCellListener mcListener = new MergedCellListener();
            sheet.VisitContainedRecords(mcListener, 0);
            int recordsAdded = mcListener.Count; 
            int recordsExpected = regionsToAdd / 1027;
            if ((regionsToAdd % 1027) != 0)
                recordsExpected++;
            Assert.AreEqual(recordsAdded, recordsExpected, "The " + regionsToAdd + " merged regions should have been spRead out over " + recordsExpected + " records, not " + recordsAdded);
            // Check we can't Add one with invalid date
            try
            {
                sheet.AddMergedRegion(10, 10, 9, 12);
                Assert.Fail("Expected an exception to occur");
            }
            catch (ArgumentException e)
            {
                // occurs during successful Test
                Assert.AreEqual("The 'to' row (9) must not be less than the 'from' row (10)", e.Message);
            }
            try
            {
                sheet.AddMergedRegion(10, 10, 12, 9);
                Assert.Fail("Expected an exception to occur");
            }
            catch (ArgumentException e)
            {
                // occurs during successful Test
                Assert.AreEqual("The 'to' col (9) must not be less than the 'from' col (10)", e.Message);
            }
        }
        [TestMethod]
        public void TestRemoveMergedRegion()
        {
            InternalSheet sheet = InternalSheet.CreateSheet();
            int regionsToAdd = 4096;

            for (int n = 0; n < regionsToAdd; n++)
                sheet.AddMergedRegion(0, 0, 1,1);

            int nSheetRecords = sheet.Records.Count;

            //remove a third from the beginning
            for (int n = 0; n < regionsToAdd / 3; n++)
            {
                sheet.RemoveMergedRegion(0);
                //assert they have been deleted
                Assert.AreEqual(sheet.NumMergedRegions,regionsToAdd - n - 1, "Num of regions should be " + (regionsToAdd - n - 1) + " not " + sheet.NumMergedRegions);
            }

            // merge records are removed from within the MergedCellsTable, 
            // so the sheet record count should not change 
            Assert.AreEqual(nSheetRecords, sheet.Records.Count, "Sheet Records");
        }

        /**
         * Bug: 22922 (Reported by Xuemin Guan)
         * 
         * Remove mergedregion Assert.Fails when a sheet loses records after an initial CreateSheet
         * fills up the records.
         *
         */
        [TestMethod]
        public void TestMovingMergedRegion()
        {
            ArrayList records = new ArrayList();

            CellRangeAddress[] cras = {
                new CellRangeAddress(0, 1, 0, 2),
            };
            MergeCellsRecord merged = new MergeCellsRecord(cras, 0, cras.Length);
            records.Add(BOFRecord.CreateSheetBOF());
            records.Add(new DimensionsRecord());
            records.Add(new RowRecord(0));
            records.Add(new RowRecord(1));
            records.Add(new RowRecord(2));
            records.Add(CreateWindow2Record());
            records.Add(EOFRecord.instance);
            records.Add(merged);

            InternalSheet sheet = CreateSheet(records);
            sheet.Records.Remove(0);

            //stub object to throw off list INDEX operations
            sheet.RemoveMergedRegion(0);
            Assert.AreEqual(0, sheet.NumMergedRegions, "Should be no more merged regions");
        }

        public void TestGetMergedRegionAt()
        {
            //TODO
        }

        public void TestGetNumMergedRegions()
        {
            //TODO
        }
        private static Record CreateWindow2Record()
        {
            WindowTwoRecord result = new WindowTwoRecord();
            result.Options=((short)0x6b6);
            result.TopRow=((short)0);
            result.LeftCol=((short)0);
            result.HeaderColor=(0x40);
            result.PageBreakZoom=((short)0);
            result.NormalZoom=((short)0);
            return result;
        }

        /**
         * Makes sure all rows registered for this sheet are aggregated, they were being skipped
         *
         */
        [TestMethod]
        public void TestRowAggregation()
        {
            ArrayList records = new ArrayList();

            records.Add(InternalSheet.CreateBOF());
            records.Add(new DimensionsRecord());
            records.Add(new RowRecord(0));
            records.Add(new RowRecord(1));
            FormulaRecord formulaRecord = new FormulaRecord();
            formulaRecord.SetCachedResultTypeString();
            records.Add(formulaRecord);
            records.Add(new StringRecord());
            records.Add(new RowRecord(2));
            records.Add(CreateWindow2Record());
            records.Add(EOFRecord.instance);

            InternalSheet sheet = CreateSheet(records);
            Assert.IsNotNull(sheet.GetRow(2), "Row [2] was skipped");
        }

        /**
         * Make sure page break functionality works (in memory)
         *
         */
        [TestMethod]
        public void TestRowPageBreaks()
        {
            short colFrom = 0;
            short colTo = 255;

            InternalSheet worksheet = InternalSheet.CreateSheet();
            PageSettingsBlock sheet = worksheet.PageSettings;
            sheet.SetRowBreak(0, colFrom, colTo);

            Assert.IsTrue(sheet.IsRowBroken(0), "no row break at 0");
            Assert.AreEqual(1, sheet.NumRowBreaks, "1 row break available");

            sheet.SetRowBreak(0, colFrom, colTo);
            sheet.SetRowBreak(0, colFrom, colTo);

            Assert.IsTrue(sheet.IsRowBroken(0), "no row break at 0");
            Assert.AreEqual(1, sheet.NumRowBreaks, "1 row break available");

            sheet.SetRowBreak(10, colFrom, colTo);
            sheet.SetRowBreak(11, colFrom, colTo);

            Assert.IsTrue(sheet.IsRowBroken(10), "no row break at 10");
            Assert.IsTrue(sheet.IsRowBroken(11), "no row break at 11");
            Assert.AreEqual(3, sheet.NumRowBreaks, "3 row break available");


            bool is10 = false;
            bool is0 = false;
            bool is11 = false;

            int[] rowBreaks = sheet.RowBreaks;
            for (int i = 0; i < rowBreaks.Length; i++)
            {
                int main = rowBreaks[i];
                if (main != 0 && main != 10 && main != 11) Assert.Fail("Invalid page break");
                if (main == 0) is0 = true;
                if (main == 10) is10 = true;
                if (main == 11) is11 = true;
            }

            Assert.IsTrue(is0 && is10 && is11, "one of the breaks didnt make it");

            sheet.RemoveRowBreak(11);
            Assert.IsFalse(sheet.IsRowBroken(11), "row should be removed");

            sheet.RemoveRowBreak(0);
            Assert.IsFalse(sheet.IsRowBroken(0), "row should be removed");

            sheet.RemoveRowBreak(10);
            Assert.IsFalse(sheet.IsRowBroken(10), "row should be removed");

            Assert.AreEqual(0, sheet.NumRowBreaks, "no more breaks");
        }

        /**
         * Make sure column pag breaks works properly (in-memory)
         *
         */
        [TestMethod]
        public void TestColPageBreaks()
        {
            int rowFrom = 0;
            int rowTo = 65535;

            InternalSheet worksheet = InternalSheet.CreateSheet();
            PageSettingsBlock sheet = worksheet.PageSettings;
            sheet.SetColumnBreak(0, rowFrom, rowTo);

            Assert.IsTrue(sheet.IsColumnBroken(0), "no col break at 0");
            Assert.AreEqual(1, sheet.NumColumnBreaks, "1 col break available");

            sheet.SetColumnBreak(0, rowFrom, rowTo);

            Assert.IsTrue(sheet.IsColumnBroken(0), "no col break at 0");
            Assert.AreEqual(1, sheet.NumColumnBreaks, "1 col break available");

            sheet.SetColumnBreak(1, rowFrom, rowTo);
            sheet.SetColumnBreak(10, rowFrom, rowTo);
            sheet.SetColumnBreak(15, rowFrom, rowTo);

            Assert.IsTrue(sheet.IsColumnBroken(1), "no col break at 1");
            Assert.IsTrue(sheet.IsColumnBroken(10), "no col break at 10");
            Assert.IsTrue(sheet.IsColumnBroken(15), "no col break at 15");
            Assert.AreEqual(4, sheet.NumColumnBreaks, "4 col break available");

            bool is10 = false;
            bool is0 = false;
            bool is1 = false;
            bool is15 = false;

            int[] colBreaks = sheet.ColumnBreaks;
            for (int i = 0; i < colBreaks.Length; i++)
            {
                int main = colBreaks[i];
                if (main != 0 && main != 1 && main != 10 && main != 15) Assert.Fail("Invalid page break");
                if (main == 0) is0 = true;
                if (main == 1) is1 = true;
                if (main == 10) is10 = true;
                if (main == 15) is15 = true;
            }

            Assert.IsTrue(is0 && is1 && is10 && is15, "one of the breaks didnt make it");

            sheet.RemoveColumnBreak(15);
            Assert.IsFalse(sheet.IsColumnBroken(15), "column break should not be there");

            sheet.RemoveColumnBreak(0);
            Assert.IsFalse(sheet.IsColumnBroken(0), "column break should not be there");

            sheet.RemoveColumnBreak(1);
            Assert.IsFalse(sheet.IsColumnBroken(1), "column break should not be there");

            sheet.RemoveColumnBreak(10);
            Assert.IsFalse(sheet.IsColumnBroken(10), "column break should not be there");

            Assert.AreEqual(0, sheet.NumColumnBreaks, "no more breaks");
        }

        /**
         * Test newly Added method Sheet.GetXFIndexForColAt(..)
         * works as designed.
         */
        [TestMethod]
        public void TestXFIndexForColumn()
        {
            short TEST_IDX = 10;
            short DEFAULT_IDX = 0xF; // 15
            short xfindex = short.MinValue;
            InternalSheet sheet = InternalSheet.CreateSheet();

            // without ColumnInfoRecord
            xfindex = sheet.GetXFIndexForColAt((short)0);
            Assert.AreEqual(DEFAULT_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)1);
            Assert.AreEqual(DEFAULT_IDX, xfindex);

            ColumnInfoRecord nci = new ColumnInfoRecord();
            sheet.ColumnInfos.InsertColumn(nci);

            // single column ColumnInfoRecord
            nci.FirstColumn = ((short)2);
            nci.LastColumn = ((short)2);
            nci.XFIndex = (TEST_IDX);
            xfindex = sheet.GetXFIndexForColAt((short)0);
            Assert.AreEqual(DEFAULT_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)1);
            Assert.AreEqual(DEFAULT_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)2);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)3);
            Assert.AreEqual(DEFAULT_IDX, xfindex);

            // ten column ColumnInfoRecord
            nci.FirstColumn = ((short)2);
            nci.LastColumn = ((short)11);
            nci.XFIndex = (TEST_IDX);
            xfindex = sheet.GetXFIndexForColAt((short)1);
            Assert.AreEqual(DEFAULT_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)2);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)6);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)11);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)12);
            Assert.AreEqual(DEFAULT_IDX, xfindex);

            // single column ColumnInfoRecord starting at index 0
            nci.FirstColumn = ((short)0);
            nci.LastColumn = ((short)0);
            nci.XFIndex = (TEST_IDX);
            xfindex = sheet.GetXFIndexForColAt((short)0);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)1);
            Assert.AreEqual(DEFAULT_IDX, xfindex);

            // ten column ColumnInfoRecord starting at index 0
            nci.FirstColumn = ((short)0);
            nci.LastColumn = ((short)9);
            nci.XFIndex = (TEST_IDX);
            xfindex = sheet.GetXFIndexForColAt((short)0);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)7);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)9);
            Assert.AreEqual(TEST_IDX, xfindex);
            xfindex = sheet.GetXFIndexForColAt((short)10);
            Assert.AreEqual(DEFAULT_IDX, xfindex);
        }
        private class SizeCheckingRecordVisitor : RecordVisitor
        {

            private int _totalSize;
            public SizeCheckingRecordVisitor()
            {
                _totalSize = 0;
            }
            public void VisitRecord(Record r)
            {

                int estimatedSize = r.RecordSize;
                byte[] buf = new byte[estimatedSize];
                int serializedSize = r.Serialize(0, buf);
                if (estimatedSize != serializedSize)
                {
                    throw new AssertFailedException("serialized size mismatch for record ("
                            + r.GetType().Name + ")");
                }
                _totalSize += estimatedSize;
            }
            public int TotalSize
            {
                get
                {
                    return _totalSize;
                }
            }
        }
        /**
         * Prior to bug 45066, POI would Get the estimated sheet size wrong
         * when an <tt>UncalcedRecord</tt> was present.<p/>
         */
        [TestMethod]
        public void TestUncalcSize_bug45066()
        {

            ArrayList records = new ArrayList();
            records.Add(new BOFRecord());
            records.Add(new UncalcedRecord());
            records.Add(new DimensionsRecord());
            records.Add(CreateWindow2Record());
            records.Add(new EOFRecord());
            InternalSheet sheet = CreateSheet(records);

            // The original bug was due to different logic for collecting records for sizing and 
            // serialization. The code has since been refactored into a single method for visiting
            // all contained records.  Now this Test is much less interesting
            SizeCheckingRecordVisitor scrv = new SizeCheckingRecordVisitor();
            sheet.VisitContainedRecords(scrv, 0);
            Assert.AreEqual(90, scrv.TotalSize);
        }

        /**
         * Prior to bug 45145 <tt>RowRecordsAggregate</tt> and <tt>ValueRecordsAggregate</tt> could
         * sometimes occur in reverse order.  This Test reproduces one of those situations and makes
         * sure that RRA comes before VRA.<br/>
         *
         * The code here represents a normal POI use case where a spReadsheet is1 Created from scratch.
         */
        [TestMethod]
        public void TestRowValueAggregatesOrder_bug45145()
        {

            InternalSheet sheet = InternalSheet.CreateSheet();

            RowRecord rr = new RowRecord(5);
            sheet.AddRow(rr);

            CellValueRecordInterface cvr = new BlankRecord();
            cvr.Column = 0;
            cvr.Row = (5);
            sheet.AddValueRecord(5, cvr);


            int dbCellRecordPos = GetDbCellRecordPos(sheet);
            if (dbCellRecordPos == 252)
            {
                // The overt symptom of the bug
                // DBCELL record pos is1 calculated wrong if VRA comes before RRA
                throw new AssertFailedException("Identified  bug 45145");
            }

            Assert.AreEqual(242, dbCellRecordPos);
        }

        /**
         * @return the value calculated for the position of the first DBCELL record for this sheet.
         * That value is1 found on the IndexRecord.
         */
        private static int GetDbCellRecordPos(InternalSheet sheet)
        {

            MyIndexRecordListener myIndexListener = new MyIndexRecordListener();
            sheet.VisitContainedRecords(myIndexListener, 0);
            IndexRecord indexRecord = myIndexListener.GetIndexRecord();
            int dbCellRecordPos = indexRecord.GetDbcellAt(0);
            return dbCellRecordPos;
        }

        private class MyIndexRecordListener : RecordVisitor
        {

            private IndexRecord _indexRecord;
            public MyIndexRecordListener()
            {
                // no-arg constructor
            }
            public IndexRecord GetIndexRecord()
            {
                return _indexRecord;
            }
            public void VisitRecord(Record r)
            {
                if (r is IndexRecord)
                {
                    if (_indexRecord != null)
                    {
                        throw new Exception("too many index records");
                    }
                    _indexRecord = (IndexRecord)r;
                }
            }
        }

        /**
 * Checks for bug introduced around r682282-r683880 that caused a second GUTS records
 * which in turn got the dimensions record out of alignment
 */
        public void TestGutsRecord_bug45640() {

            InternalSheet sheet = InternalSheet.CreateSheet();
		sheet.AddRow(new RowRecord(0));
		sheet.AddRow(new RowRecord(1));
		sheet.GroupRowRange( 0, 1, true );
		sheet.ToString();
		IList recs = sheet.Records;
		int count=0;
		for(int i=0; i< recs.Count; i++) {
			if (recs[i] is GutsRecord) {
				count++;
			}
		}
		if (count == 2) {
			throw new AssertFailedException("Identified bug 45640");
		}
		Assert.AreEqual(1, count);
	}

        public void TestMisplacedMergedCellsRecords_bug45699()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex45698-22488.xls");

            HSSFSheet sheet =(HSSFSheet)wb.GetSheetAt(0);
            HSSFRow row = (HSSFRow)sheet.GetRow(3);
            HSSFCell cell = (HSSFCell)row.GetCell(4);
            if (cell == null)
            {
                throw new AssertFailedException("Identified bug 45699");
            }
            Assert.AreEqual("Informations", cell.RichStringCellValue.String);
        }
        /**
         * In 3.1, setting margins between creating first row and first cell caused an exception.
         */
        [TestMethod]
        public void TestSetMargins_bug45717()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("Vorschauliste");
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);

            sheet.SetMargin(NPOI.SS.UserModel.MarginType.LeftMargin, 0.3);
            try
            {
                row.CreateCell(0);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("Cannot Create value records before row records exist"))
                {
                    throw new AssertFailedException("Identified bug 45717");
                }
                throw e;
            }
        }

        /**
         * Some apps seem to write files with missing DIMENSION records.
         * Excel(2007) tolerates this, so POI should too.
         */
        [TestMethod]
        public void TestMissingDims()
        {

            int rowIx = 5;
            int colIx = 6;
            NumberRecord nr = new NumberRecord();
            nr.Row = (rowIx);
            nr.Column = ((short)colIx);
            nr.Value = (3.0);

            ArrayList inRecs = new ArrayList();
            inRecs.Add(BOFRecord.CreateSheetBOF());
            inRecs.Add(new RowRecord(rowIx));
            inRecs.Add(nr);
            inRecs.Add(CreateWindow2Record());
            inRecs.Add(EOFRecord.instance);
            InternalSheet sheet;
            try
            {
                sheet = CreateSheet(inRecs);
            }
            catch (Exception e)
            {
                if ("DimensionsRecord was not found".Equals(e.Message))
                {
                    throw new AssertFailedException("Identified bug 46206");
                }
                throw e;
            }

            RecordInspector.RecordCollector rv = new RecordInspector.RecordCollector();
            sheet.VisitContainedRecords(rv, rowIx);
            Record[] outRecs = rv.Records;
            Assert.AreEqual(8, outRecs.Length);
            DimensionsRecord dims = (DimensionsRecord)outRecs[5];
            Assert.AreEqual(rowIx, dims.FirstRow);
            Assert.AreEqual(rowIx, dims.LastRow);
            Assert.AreEqual(colIx, dims.FirstCol);
            Assert.AreEqual(colIx, dims.LastCol);
        }

        /**
         * Prior to the fix for bug 46547, shifting formulas would have the side-effect
         * of creating a {@link ConditionalFormattingTable}.  There was no impairment to
         * functionality since empty record aggregates are equivalent to missing record
         * aggregates. However, since this unnecessary creation helped expose bug 46547b,
         * and since there is a slight performance hit the fix was made to avoid it.
         */
        [TestMethod]
        public void TestShiftFormulasAddCondFormat_bug46547()
        {
            // Create a sheet with data validity (similar to bugzilla attachment id=23131).
            InternalSheet sheet = InternalSheet.CreateSheet();

            IList sheetRecs = sheet.Records;
            Assert.AreEqual(24, sheetRecs.Count);

            FormulaShifter shifter = FormulaShifter.CreateForRowShift(0, 0, 0, 1);
            sheet.UpdateFormulasAfterCellShift(shifter, 0);
            if (sheetRecs.Count == 24 && sheetRecs[22] is ConditionalFormattingTable)
            {
                throw new AssertFailedException("Identified bug 46547a");
            }
            Assert.AreEqual(24, sheetRecs.Count);
        }
        /**
         * Bug 46547 happened when attempting to Add conditional formatting to a sheet
         * which already had data validity constraints.
         */
        [TestMethod]
        public void TestAddCondFormatAfterDataValidation_bug46547()
        {
            // Create a sheet with data validity (similar to bugzilla attachment id=23131).
            InternalSheet sheet = InternalSheet.CreateSheet();
            sheet.GetOrCreateDataValidityTable();

            ConditionalFormattingTable cft;
            // attempt to Add conditional formatting
            try
            {

                cft = sheet.ConditionalFormattingTable; // lazy getter
            }
            catch (InvalidCastException e)
            {
                throw new AssertFailedException("Identified bug 46547b");
            }
            Assert.IsNotNull(cft);
        }
        [TestMethod]
        public void TestCloneMulBlank_bug46776()
        {
            Record[] recs = {
				InternalSheet.CreateBOF(),
				new DimensionsRecord(),
				new RowRecord(1),
				new MulBlankRecord(1, 3, new short[] { 0x0F, 0x0F, 0x0F, } ),
				new RowRecord(2),
				CreateWindow2Record(),
				EOFRecord.instance,
		};

            InternalSheet sheet = CreateSheet(NPOI.Util.Arrays.AsList(recs));

            InternalSheet sheet2;
            try
            {
                sheet2 = sheet.CloneSheet();
            }
            catch (Exception e)
            {
                if (e.Message.Equals("The class org.apache.poi.hssf.record.MulBlankRecord needs to define a clone method"))
                {
                    throw new AssertFailedException("Identified bug 46776");
                }
                throw e;
            }

            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rc = new TestCases.HSSF.UserModel.RecordInspector.RecordCollector();
            sheet2.VisitContainedRecords(rc, 0);
            Record[] clonedRecs = rc.Records;
            Assert.AreEqual(recs.Length + 2, clonedRecs.Length); // +2 for INDEX and DBCELL
        }
    }
}

/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.HSSF.Record.Aggregates
{
    using System;
    using System.IO;
    using System.Collections;

    using NUnit.Framework;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.Util;
    using System.Collections.Generic;

    [TestFixture]
    public class TestValueRecordsAggregate
    {
        private static String ABNORMAL_SHARED_FORMULA_FLAG_TEST_FILE = "AbnormalSharedFormulaFlag.xls";
        ValueRecordsAggregate valueRecord = new ValueRecordsAggregate();
        private IList<CellValueRecordInterface> GetValueRecords()
        {
            List<CellValueRecordInterface> list = new List<CellValueRecordInterface>();
            foreach (CellValueRecordInterface rec in valueRecord)
            {
                list.Add(rec);
            }
            return list.AsReadOnly();
        }
        [TearDown]
        public void TearDown()
        {
            valueRecord = new ValueRecordsAggregate();
        }
        /**
         * Make sure the shared formula DOESNT makes it to the FormulaRecordAggregate when being parsed
         * as part of the value records
         */
        [Test]
        public void TestSharedFormula()
        {
            IList records = new ArrayList();
            records.Add(new FormulaRecord());
            records.Add(new SharedFormulaRecord());
            records.Add(new WindowTwoRecord());

            ConstructValueRecord(records);
            IList<CellValueRecordInterface> cvrs = GetValueRecords();
            //Ensure that the SharedFormulaRecord has been converted
            Assert.AreEqual(1, cvrs.Count);

            CellValueRecordInterface record = cvrs[0];
            Assert.IsNotNull(record, "Row contains a value");
            Assert.IsTrue((record is FormulaRecordAggregate), "First record is1 a FormulaRecordsAggregate");
        }

        private IList TestData()
        {
            IList records = new ArrayList();
            FormulaRecord formulaRecord = new FormulaRecord();
            //UnknownRecord unknownRecord = new UnknownRecord();
            BlankRecord blankRecord = new BlankRecord();
            WindowOneRecord windowOneRecord = new WindowOneRecord();
            formulaRecord.Row = 1;
            formulaRecord.Column = 1;
            blankRecord.Row = 2;
            blankRecord.Column = 2;
            records.Add(formulaRecord);
            records.Add(blankRecord);
            records.Add(windowOneRecord);
            return records;
        }
        [Test]
        public void TestInsertCell()
        {
            Assert.AreEqual(0, GetValueRecords().Count);

            BlankRecord blankRecord = NewBlankRecord();
            valueRecord.InsertCell(blankRecord);
            Assert.AreEqual(1, GetValueRecords().Count);
        }
        [Test]
        public void TestRemoveCell()
        {
            BlankRecord blankRecord1 = NewBlankRecord();
            valueRecord.InsertCell(blankRecord1);
            BlankRecord blankRecord2 = NewBlankRecord();
            valueRecord.RemoveCell(blankRecord2);
            Assert.AreEqual(0, GetValueRecords().Count);

            // removing an already empty cell just falls through
            valueRecord.RemoveCell(blankRecord2);

        }
        [Test]
        public void TestPhysicalNumberOfCells()
        {
            Assert.AreEqual(0, valueRecord.PhysicalNumberOfCells);
            BlankRecord blankRecord1 = NewBlankRecord();
            valueRecord.InsertCell(blankRecord1);
            Assert.AreEqual(1, valueRecord.PhysicalNumberOfCells);
            valueRecord.RemoveCell(blankRecord1);
            Assert.AreEqual(0, valueRecord.PhysicalNumberOfCells);
        }
        [Test]
        public void TestFirstCellNum()
        {
            Assert.AreEqual(-1, valueRecord.FirstCellNum);
            valueRecord.InsertCell(NewBlankRecord(2, 2));
            Assert.AreEqual(2, valueRecord.FirstCellNum);
            valueRecord.InsertCell(NewBlankRecord(3, 3));
            Assert.AreEqual(2, valueRecord.FirstCellNum);

            // Note: Removal doesn't currently reSet the first column.  It probably should but it doesn't.
            valueRecord.RemoveCell(NewBlankRecord(2, 2));
            Assert.AreEqual(2, valueRecord.FirstCellNum);
        }
        [Test]
        public void TestLastCellNum()
        {
            Assert.AreEqual(-1, valueRecord.LastCellNum);
            valueRecord.InsertCell(NewBlankRecord(2, 2));
            Assert.AreEqual(2, valueRecord.LastCellNum);
            valueRecord.InsertCell(NewBlankRecord(3, 3));
            Assert.AreEqual(3, valueRecord.LastCellNum);

            // Note: Removal doesn't currently reSet the last column.  It probably should but it doesn't.
            valueRecord.RemoveCell(NewBlankRecord(3, 3));
            Assert.AreEqual(3, valueRecord.LastCellNum);

        }
        [Test]
        public void TestSerialize()
        {
            byte[] actualArray = new byte[36];
            byte[] expectedArray = new byte[]
            {
                (byte)0x06, (byte)0x00, (byte)0x16, (byte)0x00,
                (byte)0x01, (byte)0x00, (byte)0x01, (byte)0x00,
                (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
                (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
                (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
                (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x00,
                (byte)0x00, (byte)0x00, (byte)0x01, (byte)0x02,
                (byte)0x06, (byte)0x00, (byte)0x02, (byte)0x00,
                (byte)0x02, (byte)0x00, (byte)0x00, (byte)0x00,
            };
            IList records = TestData();
            ConstructValueRecord(records);
            int bytesWritten = valueRecord.SerializeCellRow(1, 0, actualArray);
            bytesWritten += valueRecord.SerializeCellRow(2, bytesWritten, actualArray);
            Assert.AreEqual(36, bytesWritten);
            for (int i = 0; i < 36; i++)
                Assert.AreEqual(expectedArray[i], actualArray[i]);
        }

        private BlankRecord NewBlankRecord()
        {
            return NewBlankRecord(2, 2);
        }

        private BlankRecord NewBlankRecord(int col, int row)
        {
            BlankRecord blankRecord = new BlankRecord();
            blankRecord.Row = (row);
            blankRecord.Column = ((short)col);
            return blankRecord;
        }


        /**
         * Sometimes the 'shared formula' flag (<c>FormulaRecord.IsSharedFormula()</c>) is1 Set when 
         * there is1 no corresponding SharedFormulaRecord available. SharedFormulaRecord definitions do
         * not span multiple sheets.  They are are only defined within a sheet, and thus they do not 
         * have a sheet index field (only row and column range fields).<br/>
         * So it is1 important that the code which locates the SharedFormulaRecord for each 
         * FormulaRecord does not allow matches across sheets.<br/> 
         * 
         * Prior to bugzilla 44449 (Feb 2008), POI <c>ValueRecordsAggregate.construct(int, List)</c> 
         * allowed <c>SharedFormulaRecord</c>s to be erroneously used across sheets.  That incorrect
         * behaviour is1 shown by this Test.<p/>
         * 
         * <b>Notes on how to produce the Test spReadsheet</b>:
         * The Setup for this Test (AbnormalSharedFormulaFlag.xls) is1 rather fragile, insomuchas 
         * re-saving the file (either with Excel or POI) clears the flag.<br/>
         * <ol>
         * <li>A new spReadsheet was created in Excel (File | New | Blank Workbook).</li>
         * <li>Sheet3 was deleted.</li>
         * <li>Sheet2!A1 formula was Set to '="second formula"', and fill-dragged through A1:A8.</li>
         * <li>Sheet1!A1 formula was Set to '="first formula"', and also fill-dragged through A1:A8.</li>
         * <li>Four rows on Sheet1 "5" through "8" were deleted ('delete rows' alt-E D, not 'clear' Del).</li>
         * <li>The spReadsheet was saved as AbnormalSharedFormulaFlag.xls.</li>
         * </ol>
         * Prior to the row delete action the spReadsheet has two <c>SharedFormulaRecord</c>s. One 
         * for each sheet. To expose the bug, the shared formulas have been made to overlap.<br/>
         * The row delete action (as described here) seems to to delete the 
         * <c>SharedFormulaRecord</c> from Sheet1 (but not clear the 'shared formula' flags.<br/>
         * There are other variations on this theme to create the same effect.  
         * 
         */
        [Test]
        public void TestSpuriousSharedFormulaFlag()
        {

            long actualCRC = GetFileCRC(HSSFTestDataSamples.OpenSampleFileStream(ABNORMAL_SHARED_FORMULA_FLAG_TEST_FILE));
            long expectedCRC = 2277445406L;
            if (actualCRC != expectedCRC)
            {
                Console.Error.WriteLine("Expected crc " + expectedCRC + " but got " + actualCRC);
                throw failUnexpectedTestFileChange();
            }
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook(ABNORMAL_SHARED_FORMULA_FLAG_TEST_FILE);

            NPOI.SS.UserModel.ISheet s = wb.GetSheetAt(0); // Sheet1

            String cellFormula;
            cellFormula = GetFormulaFromFirstCell(s, 0); // row "1"
            // the problem is1 not observable in the first row of the shared formula
            if (!cellFormula.Equals("\"first formula\""))
            {
                throw new Exception("Something else wrong with this Test case");
            }

            // but the problem is1 observable in rows 2,3,4 
            cellFormula = GetFormulaFromFirstCell(s, 1); // row "2"
            if (cellFormula.Equals("\"second formula\""))
            {
                throw new AssertionException("found bug 44449 (Wrong SharedFormulaRecord was used).");
            }
            if (!cellFormula.Equals("\"first formula\""))
            {
                throw new Exception("Something else wrong with this Test case");
            }
        }
        private static String GetFormulaFromFirstCell(NPOI.SS.UserModel.ISheet s, int rowIx)
        {
            return s.GetRow(rowIx).GetCell((short)0).CellFormula;
        }
        private void ConstructValueRecord(IList records)
        {
            RowBlocksReader rbr = new RowBlocksReader(new RecordStream(records, 0));
            SharedValueManager sfrh = rbr.SharedFormulaManager;
            RecordStream rs = rbr.PlainRecordStream;
            while (rs.HasNext())
            {
                Record rec = rs.GetNext();
                valueRecord.Construct((CellValueRecordInterface)rec, rs, sfrh);
            }
        }
        /**
         * If someone Opened this particular Test file in Excel and saved it, the peculiar condition
         * which causes the tarGet bug would probably disappear.  This Test would then just succeed
         * regardless of whether the fix was present.  So a CRC check is1 performed to make it less easy
         * for that to occur.
         */
        private static Exception failUnexpectedTestFileChange()
        {
            String msg = "Test file '" + ABNORMAL_SHARED_FORMULA_FLAG_TEST_FILE + "' has changed.  "
                + "This junit may not be properly Testing for the tarGet bug.  "
                + "Either revert the Test file or ensure that the new version "
                + "has the right characteristics to Test the tarGet bug.";
            // A breakpoint in ValueRecordsAggregate.handleMissingSharedFormulaRecord(FormulaRecord)
            // should Get hit during parsing of Sheet1.
            // If the Test spReadsheet is1 created as directed, this condition should occur.
            // It is1 easy to upSet the Test spReadsheet (for example re-saving will destroy the 
            // peculiar condition we are Testing for). 
            throw new Exception(msg);
        }

        /**
         * Gets a CRC checksum for the content of a file
         */
        private static long GetFileCRC(Stream is1)
        {
            CRC32 crc = new CRC32();
            return crc.StreamCRC(is1);
        }
        [Test]
        public void TestRemoveNewRow_bug46312()
        {
            // To make bug occur, rowIndex needs to be >= ValueRecordsAggregate.records.length
            int rowIndex = 30;

            ValueRecordsAggregate vra = new ValueRecordsAggregate();
            try
            {
                vra.RemoveAllCellsValuesForRow(rowIndex);
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Specified rowIndex 30 is outside the allowable range (0..30)"))
                {
                    throw new AssertionException("Identified bug 46312");
                }
                throw e;
            }

            //if (false) { // same bug as demonstrated through usermodel API

            //    HSSFWorkbook wb = new HSSFWorkbook();
            //    HSSFSheet sheet = (HSSFSheet)wb.CreateSheet();
            //    HSSFRow row = (HSSFRow)sheet.CreateRow(rowIndex);
            //    if (false) { // must not add any cells to the new row if we want to see the bug
            //        row.CreateCell(0); // this causes ValueRecordsAggregate.records to auto-extend
            //    }
            //    try {
            //        sheet.CreateRow(rowIndex);
            //    } catch (ArgumentException e) {
            //        throw new AssertionException("Identified bug 46312");
            //    }
            //}
        }

        /**
         * Tests various manipulations of blank cells, to make sure that {@link MulBlankRecord}s
         * are use appropriately
         */
        [Test]
        public void TestMultipleBlanks()
        {
            BlankRecord brA2 = NewBlankRecord(0, 1);
            BlankRecord brB2 = NewBlankRecord(1, 1);
            BlankRecord brC2 = NewBlankRecord(2, 1);
            BlankRecord brD2 = NewBlankRecord(3, 1);
            BlankRecord brE2 = NewBlankRecord(4, 1);
            BlankRecord brB3 = NewBlankRecord(1, 2);
            BlankRecord brC3 = NewBlankRecord(2, 2);

            valueRecord.InsertCell(brA2);
            valueRecord.InsertCell(brB2);
            valueRecord.InsertCell(brD2);
            confirmMulBlank(3, 1, 1);

            valueRecord.InsertCell(brC3);
            confirmMulBlank(4, 1, 2);

            valueRecord.InsertCell(brB3);
            valueRecord.InsertCell(brE2);
            confirmMulBlank(6, 3, 0);

            valueRecord.InsertCell(brC2);
            confirmMulBlank(7, 2, 0);

            valueRecord.RemoveCell(brA2);
            confirmMulBlank(6, 2, 0);

            valueRecord.RemoveCell(brC2);
            confirmMulBlank(5, 2, 1);

            valueRecord.RemoveCell(brC3);
            confirmMulBlank(4, 1, 2);
        }
        private class RecordVisitor1 : RecordVisitor
        {
            private BlankStats bs;
            public RecordVisitor1(BlankStats bs)
            {
                this.bs = bs;
            }
            #region RecordVisitor ≥…‘±

            public void VisitRecord(Record r)
            {
                if (r is MulBlankRecord)
                {
                    MulBlankRecord mbr = (MulBlankRecord)r;
                    bs.countMulBlankRecords++;
                    bs.countBlankCells += mbr.NumColumns;
                }
                else if (r is BlankRecord)
                {
                    bs.countSingleBlankRecords++;
                    bs.countBlankCells++;
                }
            }

            #endregion
        }
        class BlankStats
        {
            public int countBlankCells;
            public int countMulBlankRecords;
            public int countSingleBlankRecords;
        }
        private void confirmMulBlank(int expectedTotalBlankCells,
                int expectedNumberOfMulBlankRecords, int expectedNumberOfSingleBlankRecords)
        {
            // assumed row ranges set-up by caller:
            int firstRow = 1;
            int lastRow = 2;
            BlankStats bs = new BlankStats();
            RecordVisitor rv = new RecordVisitor1(bs);

            for (int rowIx = firstRow; rowIx <= lastRow; rowIx++)
            {
                if (valueRecord.RowHasCells(rowIx))
                {
                    valueRecord.VisitCellsForRow(rowIx, rv);
                }
            }
            Assert.AreEqual(expectedTotalBlankCells, bs.countBlankCells);
            Assert.AreEqual(expectedNumberOfMulBlankRecords, bs.countMulBlankRecords);
            Assert.AreEqual(expectedNumberOfSingleBlankRecords, bs.countSingleBlankRecords);
        }
    }
}

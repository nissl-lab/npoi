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

namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;

    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using TestCases.SS.UserModel;
    using NPOI.SS;
    using NPOI.HSSF.Record;

    /**
     * Test Row is okay.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestHSSFRow : BaseTestRow
    {
        public TestHSSFRow(): base(HSSFITestDataProvider.Instance)
        {
          
        }

        [Test]
        public void TestMoveCell()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet();
            IRow row = sheet.CreateRow(0);
            IRow rowB = sheet.CreateRow(1);

            ICell cellA2 = rowB.CreateCell(0);
            Assert.AreEqual(0, rowB.FirstCellNum);
            Assert.AreEqual(0, rowB.FirstCellNum);

            Assert.AreEqual(-1, row.LastCellNum);
            Assert.AreEqual(-1, row.FirstCellNum);
            ICell cellB2 = row.CreateCell(1);
            ICell cellB3 = row.CreateCell(2);
            ICell cellB4 = row.CreateCell(3);

            Assert.AreEqual(1, row.FirstCellNum);
            Assert.AreEqual(4, row.LastCellNum);

            // Try to move to somewhere else that's used
            try
            {
                row.MoveCell(cellB2, (short)3);
                Assert.Fail("ArgumentException should have been thrown");
            }
            catch (ArgumentException)
            {
                // expected during successful Test
            }

            // Try to move one off a different row
            try
            {
                row.MoveCell(cellA2, (short)3);
                Assert.Fail("ArgumentException should have been thrown");
            }
            catch (ArgumentException)
            {
                // expected during successful Test
            }

            // Move somewhere spare
            Assert.IsNotNull(row.GetCell(1));
            row.MoveCell(cellB2, (short)5);
            Assert.IsNull(row.GetCell(1));
            Assert.IsNotNull(row.GetCell(5));

            Assert.AreEqual(5, cellB2.ColumnIndex);
            Assert.AreEqual(2, row.FirstCellNum);
            Assert.AreEqual(6, row.LastCellNum);
        }
        [Test]
        public void TestRowBounds()
        {
            BaseTestRowBounds(SpreadsheetVersion.EXCEL97.LastRowIndex);
        }
        [Test]
        public void TestCellBounds()
        {
            BaseTestCellBounds(SpreadsheetVersion.EXCEL97.LastColumnIndex);
        }
        /**
         * Prior to patch 43901, POI was producing files with the wrong last-column
         * number on the row
         */
        [Test]
        public new void TestLastCellNumIsCorrectAfterAddCell_bug43901()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = book.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);

            // New row has last col -1
            Assert.AreEqual(-1, row.LastCellNum);
            if (row.LastCellNum == 0)
            {
                Assert.Fail("Identified bug 43901");
            }

            // Create two cells, will return one higher
            //  than that for the last number
            row.CreateCell(0);
            Assert.AreEqual(1, row.LastCellNum);
            row.CreateCell(255);
            Assert.AreEqual(256, row.LastCellNum);
        }
        [Test]
        public void TestLastAndFirstColumns_bug46654()
        {
            int ROW_IX = 10;
            int COL_IX = 3;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("Sheet1");
            RowRecord rowRec = new RowRecord(ROW_IX);
            rowRec.FirstCol=((short)2);
            rowRec.LastCol=((short)5);

            BlankRecord br = new BlankRecord();
            br.Row=(ROW_IX);
            br.Column=((short)COL_IX);

            sheet.Sheet.AddValueRecord(ROW_IX, br);
            HSSFRow row = new HSSFRow(workbook,sheet, rowRec);
            ICell cell = row.CreateCellFromRecord(br);

            if (row.FirstCellNum == 2 && row.LastCellNum == 5)
            {
                throw new AssertionException("Identified bug 46654a");
            }
            Assert.AreEqual(COL_IX, row.FirstCellNum);
            Assert.AreEqual(COL_IX + 1, row.LastCellNum);
            row.RemoveCell(cell);
            Assert.AreEqual(-1, row.FirstCellNum);
            Assert.AreEqual(-1, row.LastCellNum);
        }

        [Test]
        public new void TestRowHeight()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;

            Assert.AreEqual(row.Height, sheet.DefaultRowHeight);
            Assert.AreEqual(row.RowRecord.BadFontHeight, false);

            row.Height=((short)123);
            Assert.AreEqual(row.Height, 123);
            Assert.AreEqual(row.RowRecord.BadFontHeight, true);

            row.Height = ((short)-1);
            Assert.AreEqual(row.Height, sheet.DefaultRowHeight);
            Assert.AreEqual(row.RowRecord.BadFontHeight, false);
        }
    }
}
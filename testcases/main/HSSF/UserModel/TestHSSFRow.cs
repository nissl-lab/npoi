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
    using NUnit.Framework;using NUnit.Framework.Legacy;

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
        public void TestRowBounds()
        {
            BaseTestRowBounds(SpreadsheetVersion.EXCEL97.LastRowIndex);
        }
        [Test]
        public void TestCellBounds()
        {
            BaseTestCellBounds(SpreadsheetVersion.EXCEL97.LastColumnIndex);
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
                Assert.Fail("Identified bug 46654a");
            }
            ClassicAssert.AreEqual(COL_IX, row.FirstCellNum);
            ClassicAssert.AreEqual(COL_IX + 1, row.LastCellNum);
            row.RemoveCell(cell);
            ClassicAssert.AreEqual(-1, row.FirstCellNum);
            ClassicAssert.AreEqual(-1, row.LastCellNum);
            workbook.Close();
        }

        [Test]
        public void TestMoveCell()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row = sheet.CreateRow(0);
            IRow rowB = sheet.CreateRow(1);

            ICell cellA2 = rowB.CreateCell(0);
            ClassicAssert.AreEqual(0, rowB.FirstCellNum);
            ClassicAssert.AreEqual(0, rowB.FirstCellNum);

            ClassicAssert.AreEqual(-1, row.LastCellNum);
            ClassicAssert.AreEqual(-1, row.FirstCellNum);
            ICell cellB2 = row.CreateCell(1);
            ICell cellB3 = row.CreateCell(2);
            ICell cellB4 = row.CreateCell(3);

            ClassicAssert.AreEqual(1, row.FirstCellNum);
            ClassicAssert.AreEqual(4, row.LastCellNum);

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
            ClassicAssert.IsNotNull(row.GetCell(1));
            row.MoveCell(cellB2, (short)5);
            ClassicAssert.IsNull(row.GetCell(1));
            ClassicAssert.IsNotNull(row.GetCell(5));

            ClassicAssert.AreEqual(5, cellB2.ColumnIndex);
            ClassicAssert.AreEqual(2, row.FirstCellNum);
            ClassicAssert.AreEqual(6, row.LastCellNum);

            workbook.Close();
        }

        [Test]
        public new void TestRowHeight()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;

            ClassicAssert.AreEqual(row.Height, sheet.DefaultRowHeight);
            ClassicAssert.IsFalse(row.RowRecord.BadFontHeight);

            row.Height=((short)123);
            ClassicAssert.AreEqual(123, row.Height);
            ClassicAssert.IsTrue(row.RowRecord.BadFontHeight);

            row.Height = ((short)-1);
            ClassicAssert.AreEqual(row.Height, sheet.DefaultRowHeight);
            ClassicAssert.IsFalse(row.RowRecord.BadFontHeight);

            row.Height=((short) 123);
            ClassicAssert.AreEqual(123, row.Height);
            ClassicAssert.IsTrue(row.RowRecord.BadFontHeight);

            row.HeightInPoints = (-1);
            ClassicAssert.AreEqual(row.Height, sheet.DefaultRowHeight);
            ClassicAssert.IsFalse(row.RowRecord.BadFontHeight);

            row.HeightInPoints = (432);
            ClassicAssert.AreEqual(432*20, row.Height);
            ClassicAssert.IsTrue(row.RowRecord.BadFontHeight);

            workbook.Close();
        }
    }
}
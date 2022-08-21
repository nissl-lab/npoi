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

namespace TestCases.SS.Util
{
    using System;

    using NUnit.Framework;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * Tests SheetUtil.
     *
     * @see NPOI.SS.Util.SheetUtil
     */
    [TestFixture]
    public class TestSheetUtil
    {
        [Test]
        public void TestCellWithMerges()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();

            // Create some test data
            IRow r2 = s.CreateRow(1);
            r2.CreateCell(0).SetCellValue(10);
            r2.CreateCell(1).SetCellValue(11);
            IRow r3 = s.CreateRow(2);
            r3.CreateCell(0).SetCellValue(20);
            r3.CreateCell(1).SetCellValue(21);

            s.AddMergedRegion(new CellRangeAddress(2, 3, 0, 0));
            s.AddMergedRegion(new CellRangeAddress(2, 2, 1, 4));

            // With a cell that isn't defined, we'll Get null
            Assert.AreEqual(null, SheetUtil.GetCellWithMerges(s, 0, 0));

            // With a cell that's not in a merged region, we'll Get that
            Assert.AreEqual(10.0, SheetUtil.GetCellWithMerges(s, 1, 0).NumericCellValue);
            Assert.AreEqual(11.0, SheetUtil.GetCellWithMerges(s, 1, 1).NumericCellValue);

            // With a cell that's the primary one of a merged region, we Get that cell
            Assert.AreEqual(20.0, SheetUtil.GetCellWithMerges(s, 2, 0).NumericCellValue);
            Assert.AreEqual(21, SheetUtil.GetCellWithMerges(s, 2, 1).NumericCellValue);

            // With a cell elsewhere in the merged region, Get top-left
            Assert.AreEqual(20.0, SheetUtil.GetCellWithMerges(s, 3, 0).NumericCellValue);
            Assert.AreEqual(21.0, SheetUtil.GetCellWithMerges(s, 2, 2).NumericCellValue);
            Assert.AreEqual(21.0, SheetUtil.GetCellWithMerges(s, 2, 3).NumericCellValue);
            Assert.AreEqual(21.0, SheetUtil.GetCellWithMerges(s, 2, 4).NumericCellValue);

            wb.Close();
        }
        [Test]
        public void testCanComputeWidthHSSF()
        {
            IWorkbook wb = new HSSFWorkbook();

            // cannot check on result because on some machines we get back false here!
            SheetUtil.CanComputeColumnWidth(wb.GetFontAt((short)0));
            wb.Close();
        }
        [Test]
        public void testGetCellWidthEmpty()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            // no contents: cell.setCellValue("sometext");

            Assert.AreEqual(-1.0, SheetUtil.GetCellWidth(cell, 1, null, true));

            wb.Close();
        }
        [Test]
        public void testGetCellWidthString()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            cell.SetCellValue("sometext");

            Assert.IsTrue(SheetUtil.GetCellWidth(cell, 1, null, true) > 0);

            wb.Close();
        }
        [Test]
        public void testGetCellWidthNumber()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            cell.SetCellValue(88.234);

            Assert.IsTrue(SheetUtil.GetCellWidth(cell, 1, null, true) > 0);

            wb.Close();
        }
        [Test]
        public void testGetCellWidthBoolean()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            cell.SetCellValue(false);

            Assert.IsTrue(SheetUtil.GetCellWidth(cell, 1, null, false) > 0);

            wb.Close();
        }
        [Test]
        public void testGetColumnWidthString()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet");
            IRow row = sheet.CreateRow(0);
            sheet.CreateRow(1);
            sheet.CreateRow(2);
            ICell cell = row.CreateCell(0);

            cell.SetCellValue("sometext");

            Assert.IsTrue(SheetUtil.GetColumnWidth(sheet, 0, true) > 0, "Having some width for rows with actual cells");
            Assert.AreEqual(-1.0, SheetUtil.GetColumnWidth(sheet, 0, true, 1, 2)
                    , "Not having any widht for rows with all empty cells");

            wb.Close();
        }

        [Test]
        public void testGetRowHeight()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            row.CreateCell(1);
            row.CreateCell(2);

            cell.SetCellValue("sometext");
            
            Assert.IsTrue(SheetUtil.GetRowHeight(sheet, 0, true) > 0, "Having some height for a row with a cell with content");
            Assert.AreEqual(0.0, SheetUtil.GetRowHeight(sheet, 0, true, 1, 2)
                    , "Not having any height for row with empty cells");
            Assert.IsTrue(SheetUtil.GetRowHeight(row, true) > 0, "Having some height for a row with a cell with content");
            Assert.AreEqual(0.0, SheetUtil.GetRowHeight(row, true, 1, 2)
                    , "Not having any height for row with empty cells");

            wb.Close();
        }


        [Test]
        public void testCellHeight()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            ICell emptyCell = row.CreateCell(1);

            cell.SetCellValue("sometext");

            Assert.IsTrue(SheetUtil.GetCellHeight(cell, false) > 0, "Having some height for a cell with content");
            Assert.AreEqual(0.0, SheetUtil.GetCellHeight(emptyCell, false), "Not having any height for a cell with no content");

            wb.Close();
        }
    }
}
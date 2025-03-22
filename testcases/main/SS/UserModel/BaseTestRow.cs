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

namespace TestCases.SS.UserModel
{
    using System;



    using NUnit.Framework;using NUnit.Framework.Legacy;

    using NPOI.SS;
    using System.Collections;
    using TestCases.SS;
    using NPOI.SS.UserModel;

    /**
     * A base class for Testing implementations of
     * {@link NPOI.SS.UserModel.Row}
     */
    public class BaseTestRow
    {

        protected ITestDataProvider _testDataProvider;
        public BaseTestRow()
            : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        { }
        protected BaseTestRow(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }
        [Test]
        public void TestLastAndFirstColumns()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ClassicAssert.AreEqual(-1, row.FirstCellNum);
            ClassicAssert.AreEqual(-1, row.LastCellNum);

            //getting cells from an empty row should returns null
            for (int i = 0; i < 10; i++) ClassicAssert.IsNull(row.GetCell(i));

            row.CreateCell(2);
            ClassicAssert.AreEqual(2, row.FirstCellNum);
            ClassicAssert.AreEqual(3, row.LastCellNum);

            row.CreateCell(1);
            ClassicAssert.AreEqual(1, row.FirstCellNum);
            ClassicAssert.AreEqual(3, row.LastCellNum);

            // check the exact case reported in 'bug' 43901 - notice that the cellNum is '0' based
            row.CreateCell(3);
            ClassicAssert.AreEqual(1, row.FirstCellNum);
            ClassicAssert.AreEqual(4, row.LastCellNum);

            workbook.Close();
        }

        /**
         * Make sure that there is no cross-talk between rows especially with GetFirstCellNum and GetLastCellNum
         * This Test was Added in response to bug report 44987.
         */
        [Test]
        public void TestBoundsInMultipleRows()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow rowA = sheet.CreateRow(0);

            rowA.CreateCell(10);
            rowA.CreateCell(5);
            ClassicAssert.AreEqual(5, rowA.FirstCellNum);
            ClassicAssert.AreEqual(11, rowA.LastCellNum);

            IRow rowB = sheet.CreateRow(1);
            rowB.CreateCell(15);
            rowB.CreateCell(30);
            ClassicAssert.AreEqual(15, rowB.FirstCellNum);
            ClassicAssert.AreEqual(31, rowB.LastCellNum);

            ClassicAssert.AreEqual(5, rowA.FirstCellNum);
            ClassicAssert.AreEqual(11, rowA.LastCellNum);
            rowA.CreateCell(50);
            ClassicAssert.AreEqual(51, rowA.LastCellNum);

            ClassicAssert.AreEqual(31, rowB.LastCellNum);

            workbook.Close();
        }
        [Test]
        public void TestRemoveCell()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();
            IRow row = sheet.CreateRow(0);

            ClassicAssert.AreEqual(0, row.PhysicalNumberOfCells);
            ClassicAssert.AreEqual(-1, row.LastCellNum);
            ClassicAssert.AreEqual(-1, row.FirstCellNum);

            row.CreateCell(1);
            ClassicAssert.AreEqual(2, row.LastCellNum);
            ClassicAssert.AreEqual(1, row.FirstCellNum);
            ClassicAssert.AreEqual(1, row.PhysicalNumberOfCells);
            row.CreateCell(3);
            ClassicAssert.AreEqual(4, row.LastCellNum);
            ClassicAssert.AreEqual(1, row.FirstCellNum);
            ClassicAssert.AreEqual(2, row.PhysicalNumberOfCells);
            row.RemoveCell(row.GetCell(3));
            ClassicAssert.AreEqual(2, row.LastCellNum);
            ClassicAssert.AreEqual(1, row.FirstCellNum);
            ClassicAssert.AreEqual(1, row.PhysicalNumberOfCells);
            row.RemoveCell(row.GetCell(1));
            ClassicAssert.AreEqual(-1, row.LastCellNum);
            ClassicAssert.AreEqual(-1, row.FirstCellNum);
            ClassicAssert.AreEqual(0, row.PhysicalNumberOfCells);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            row = sheet.GetRow(0);
            ClassicAssert.AreEqual(-1, row.LastCellNum);
            ClassicAssert.AreEqual(-1, row.FirstCellNum);
            ClassicAssert.AreEqual(0, row.PhysicalNumberOfCells);
            wb2.Close();
        }
        protected void BaseTestRowBounds(int maxRowNum)
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            //Test low row bound
            sheet.CreateRow(0);
            //Test low row bound exception
            try
            {
                sheet.CreateRow(-1);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                ClassicAssert.IsTrue(e.Message.StartsWith("Invalid row number (-1)"),
                    "Did not find expected error message, had: " + e);
            }

            //Test high row bound
            sheet.CreateRow(maxRowNum);
            //Test high row bound exception
            try
            {
                sheet.CreateRow(maxRowNum + 1);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                ClassicAssert.AreEqual("Invalid row number (" + (maxRowNum + 1) + ") outside allowable range (0.." + maxRowNum + ")", e.Message);
            }
            workbook.Close();
        }
        protected void BaseTestCellBounds(int maxCellNum)
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();

            IRow row = sheet.CreateRow(0);
            //Test low cell bound
            try
            {
                row.CreateCell(-1);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                ClassicAssert.IsTrue(e.Message.StartsWith("Invalid column index (-1)"));
            }

            //Test high cell bound
            try
            {
                row.CreateCell(maxCellNum + 1);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                ClassicAssert.IsTrue(e.Message.StartsWith("Invalid column index (" + (maxCellNum + 1) + ")"));
            }
            for (int i = 0; i < maxCellNum; i++)
            {
                row.CreateCell(i);
            }
            ClassicAssert.AreEqual(maxCellNum, row.PhysicalNumberOfCells);
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            row = sheet.GetRow(0);
            ClassicAssert.AreEqual(maxCellNum, row.PhysicalNumberOfCells);
            for (int i = 0; i < maxCellNum; i++)
            {
                ICell cell = row.GetCell(i);
                ClassicAssert.AreEqual(i, cell.ColumnIndex);
            }
            wb2.Close();
        }

        /**
         * Prior to patch 43901, POI was producing files with the wrong last-column
         * number on the row
         */
        [Test]
        public void TestLastCellNumIsCorrectAfterAddCell_bug43901()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("test");
            IRow row = sheet.CreateRow(0);

            // New row has last col -1
            ClassicAssert.AreEqual(-1, row.LastCellNum);
            if (row.LastCellNum == 0)
            {
                Assert.Fail("Identified bug 43901");
            }

            // Create two cells, will return one higher
            //  than that for the last number
            row.CreateCell(0);
            ClassicAssert.AreEqual(1, row.LastCellNum);
            row.CreateCell(255);
            ClassicAssert.AreEqual(256, row.LastCellNum);

            workbook.Close();
        }

        /**
         * Tests for the missing/blank cell policy stuff
         */
        [Test]
        public void TestGetCellPolicy()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("test");
            IRow row = sheet.CreateRow(0);

            // 0 -> string
            // 1 -> num
            // 2 missing
            // 3 missing
            // 4 -> blank
            // 5 -> num
            row.CreateCell(0).SetCellValue("test");
            row.CreateCell(1).SetCellValue(3.2);
            row.CreateCell(4, CellType.Blank);
            row.CreateCell(5).SetCellValue(4);

            // First up, no policy given, uses default
            ClassicAssert.AreEqual(CellType.String, row.GetCell(0).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(1).CellType);
            ClassicAssert.AreEqual(null, row.GetCell(2));
            ClassicAssert.AreEqual(null, row.GetCell(3));
            ClassicAssert.AreEqual(CellType.Blank, row.GetCell(4).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(5).CellType);

            // RETURN_NULL_AND_BLANK - same as default
            ClassicAssert.AreEqual(CellType.String, row.GetCell(0, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(1, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);
            ClassicAssert.AreEqual(null, row.GetCell(2, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            ClassicAssert.AreEqual(null, row.GetCell(3, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            ClassicAssert.AreEqual(CellType.Blank, row.GetCell(4, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(5, MissingCellPolicy.RETURN_NULL_AND_BLANK).CellType);

            // RETURN_BLANK_AS_NULL - nearly the same
            ClassicAssert.AreEqual(CellType.String, row.GetCell(0, MissingCellPolicy.RETURN_BLANK_AS_NULL).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(1, MissingCellPolicy.RETURN_BLANK_AS_NULL).CellType);
            ClassicAssert.AreEqual(null, row.GetCell(2, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            ClassicAssert.AreEqual(null, row.GetCell(3, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            ClassicAssert.AreEqual(null, row.GetCell(4, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(5, MissingCellPolicy.RETURN_BLANK_AS_NULL).CellType);

            // CREATE_NULL_AS_BLANK - Creates as needed
            ClassicAssert.AreEqual(CellType.String, row.GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            ClassicAssert.AreEqual(CellType.Blank, row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            ClassicAssert.AreEqual(CellType.Blank, row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            ClassicAssert.AreEqual(CellType.Blank, row.GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).CellType);

            // Check Created ones Get the right column
            ClassicAssert.AreEqual(0, row.GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            ClassicAssert.AreEqual(1, row.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            ClassicAssert.AreEqual(2, row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            ClassicAssert.AreEqual(3, row.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            ClassicAssert.AreEqual(4, row.GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);
            ClassicAssert.AreEqual(5, row.GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).ColumnIndex);


            // Now change the cell policy on the workbook, check
            //  that that is now used if no policy given
            workbook.MissingCellPolicy = (MissingCellPolicy.RETURN_BLANK_AS_NULL);

            ClassicAssert.AreEqual(CellType.String, row.GetCell(0).CellType);
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(1).CellType);
            ClassicAssert.AreEqual(null, row.GetCell(2));
            ClassicAssert.AreEqual(null, row.GetCell(3));
            ClassicAssert.AreEqual(null, row.GetCell(4));
            ClassicAssert.AreEqual(CellType.Numeric, row.GetCell(5).CellType);

            workbook.Close();
        }
        [Test]
        public void TestRowHeight()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();
            IRow row1 = sheet.CreateRow(0);

            ClassicAssert.AreEqual(sheet.DefaultRowHeight, row1.Height);

            sheet.DefaultRowHeightInPoints = (/*setter*/20);
            row1.Height = (short)-1; //reset the row height
            ClassicAssert.AreEqual(20.0f, row1.HeightInPoints, 0F);
            ClassicAssert.AreEqual(20 * 20, row1.Height);

            IRow row2 = sheet.CreateRow(1);
            ClassicAssert.AreEqual(sheet.DefaultRowHeight, row2.Height);
            row2.Height = (short)310;
            ClassicAssert.AreEqual(310, row2.Height);
            ClassicAssert.AreEqual(310F / 20, row2.HeightInPoints, 0F);

            IRow row3 = sheet.CreateRow(2);
            row3.HeightInPoints = (/*setter*/25.5f);
            ClassicAssert.AreEqual((short)(25.5f * 20), row3.Height);
            ClassicAssert.AreEqual(25.5f, row3.HeightInPoints, 0F);

            IRow row4 = sheet.CreateRow(3);
            ClassicAssert.IsFalse(row4.ZeroHeight);
            row4.ZeroHeight = (/*setter*/true);
            ClassicAssert.IsTrue(row4.ZeroHeight);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            sheet = wb2.GetSheetAt(0);

            row1 = sheet.GetRow(0);
            row2 = sheet.GetRow(1);
            row3 = sheet.GetRow(2);
            row4 = sheet.GetRow(3);
            ClassicAssert.AreEqual(20.0f, row1.HeightInPoints, 0F);
            ClassicAssert.AreEqual(20 * 20, row1.Height);

            ClassicAssert.AreEqual(310, row2.Height);
            ClassicAssert.AreEqual(310F / 20, row2.HeightInPoints, 0F);

            ClassicAssert.AreEqual((short)(25.5f * 20), row3.Height);
            ClassicAssert.AreEqual(25.5f, row3.HeightInPoints, 0F);

            ClassicAssert.IsFalse(row1.ZeroHeight);
            ClassicAssert.IsFalse(row2.ZeroHeight);
            ClassicAssert.IsFalse(row3.ZeroHeight);
            ClassicAssert.IsTrue(row4.ZeroHeight);

            wb2.Close();
        }

        /**
         * Test Adding cells to a row in various places and see if we can find them again.
         */
        [Test]
        public void TestCellIterator()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);

            // One cell at the beginning
            ICell cell1 = row.CreateCell(1);
            IEnumerator it = row.GetEnumerator();
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell1 == it.Current);
            ClassicAssert.IsFalse(it.MoveNext());

            // Add another cell at the end
            ICell cell2 = row.CreateCell(99);
            it = row.GetEnumerator();
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell1 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell2 == it.Current);

            // Add another cell at the beginning
            ICell cell3 = row.CreateCell(0);
            it = row.GetEnumerator();
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell3 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell1 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell2 == it.Current);

            // Replace cell1
            ICell cell4 = row.CreateCell(1);
            it = row.GetEnumerator();
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell3 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell4 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell2 == it.Current);
            ClassicAssert.IsFalse(it.MoveNext());

            // Add another cell, specifying the cellType
            ICell cell5 = row.CreateCell(2, CellType.String);
            it = row.GetEnumerator();
            ClassicAssert.IsNotNull(cell5);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell3 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell4 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell5 == it.Current);
            ClassicAssert.IsTrue(it.MoveNext());
            ClassicAssert.IsTrue(cell2 == it.Current);
            ClassicAssert.AreEqual(CellType.String, cell5.CellType);

            wb.Close();
        }
        [Test]
        public void TestRowStyle()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet("test");
            IRow row1 = sheet.CreateRow(0);
            IRow row2 = sheet.CreateRow(1);

            // Won't be styled currently
            ClassicAssert.AreEqual(false, row1.IsFormatted);
            ClassicAssert.AreEqual(false, row2.IsFormatted);
            ClassicAssert.AreEqual(null, row1.RowStyle);
            ClassicAssert.AreEqual(null, row2.RowStyle);

            // Style one
            ICellStyle style = wb1.CreateCellStyle();
            style.DataFormat = (/*setter*/(short)4);
            row2.RowStyle = (/*setter*/style);

            // Check
            ClassicAssert.AreEqual(false, row1.IsFormatted);
            ClassicAssert.AreEqual(true, row2.IsFormatted);
            ClassicAssert.AreEqual(null, row1.RowStyle);
            ClassicAssert.AreEqual(style, row2.RowStyle);

            // Save, load and re-check
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            sheet = wb2.GetSheetAt(0);

            row1 = sheet.GetRow(0);
            row2 = sheet.GetRow(1);
            style = wb2.GetCellStyleAt(style.Index);

            ClassicAssert.AreEqual(false, row1.IsFormatted);
            ClassicAssert.AreEqual(true, row2.IsFormatted);
            ClassicAssert.AreEqual(null, row1.RowStyle);
            ClassicAssert.AreEqual(style, row2.RowStyle);
            ClassicAssert.AreEqual(4, style.DataFormat);

            wb2.Close();
        }
    }

}
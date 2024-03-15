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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using TestCases.SS.UserModel;
namespace TestCases.XSSF.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFSheetShiftRows : BaseTestSheetShiftRows
    {

        public TestXSSFSheetShiftRows()
            : base(XSSFITestDataProvider.instance)
        {

        }

        [Test]
        public override void TestShiftRowBreaks()
        {
            // disabled test from superclass
            // TODO - support shifting of page breaks
        }

        [Test]
        public void TestBug54524()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("54524.xlsx");
            ISheet sheet = workbook.GetSheetAt(0);
            sheet.ShiftRows(3, 5, -1);

            ICell cell = CellUtil.GetCell(sheet.GetRow(1), 0);
            Assert.AreEqual(1.0, cell.NumericCellValue, 0);
            cell = CellUtil.GetCell(sheet.GetRow(2), 0);
            Assert.AreEqual("SUM(A2:A2)", cell.CellFormula);
            cell = CellUtil.GetCell(sheet.GetRow(3), 0);
            Assert.AreEqual("X", cell.StringCellValue);

            workbook.Close();
        }
        [Test]
        public void TestBug53798()
        {
            // NOTE that for HSSF (.xls) negative shifts combined with positive ones do work as expected  
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53798.xlsx");

            ISheet testSheet = wb.GetSheetAt(0);
            // 1) corrupted xlsx (unreadable data in the first row of a shifted group) already comes about  
            // when shifted by less than -1 negative amount (try -2)
            testSheet.ShiftRows(3, 3, -2);

            IRow newRow = null; ICell newCell = null;
            // 2) attempt to create a new row IN PLACE of a removed row by a negative shift causes corrupted 
            // xlsx file with  unreadable data in the negative shifted row. 
            // NOTE it's ok to create any other row.
            newRow = testSheet.CreateRow(3);
            newCell = newRow.CreateCell(0);
            newCell.SetCellValue("new Cell in row " + newRow.RowNum);

            // 3) once a negative shift has been made any attempt to shift another group of rows 
            // (note: outside of previously negative shifted rows) by a POSITIVE amount causes POI exception: 
            // org.apache.xmlbeans.impl.values.XmlValueDisconnectedException.
            // NOTE: another negative shift on another group of rows is successful, provided no new rows in  
            // place of previously shifted rows were attempted to be created as explained above.
            // -- CHANGE the shift to positive once the behaviour of the above has been tested
            testSheet.ShiftRows(6, 7, 1);
            

            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            wb.Close();
            Assert.IsNotNull(read);

            ISheet readSheet = read.GetSheetAt(0);
            verifyCellContent(readSheet, 0, "0.0");
            verifyCellContent(readSheet, 1, "3.0");
            verifyCellContent(readSheet, 2, "2.0");
            verifyCellContent(readSheet, 3, "new Cell in row 3");
            verifyCellContent(readSheet, 4, "4.0");
            verifyCellContent(readSheet, 5, "5.0");
            verifyCellContent(readSheet, 6, null);
            verifyCellContent(readSheet, 7, "6.0");
            verifyCellContent(readSheet, 8, "7.0");
            read.Close();
        }

        // CT_Rows should stay sorted in ascending order after a call to ShiftRows.
        [Test]
        public void TestBug57423()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet wsh = sheet.GetCTWorksheet();
            CT_SheetData sheetData = wsh.sheetData;

            XSSFRow row1 = (XSSFRow)sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue("a");

            XSSFRow row2 = (XSSFRow) sheet.CreateRow(1);
            row2.CreateCell(0).SetCellValue("b");

            XSSFRow row3 = (XSSFRow) sheet.CreateRow(2);
            row3.CreateCell(0).SetCellValue("c");

            sheet.ShiftRows(0, 1, 3); //move "a" and "b" 3 rows down
            //      Before:    After:
            //         A        A
            // 1       a        <empty>
            // 2       b        <empty>
            // 3       c        c
            // 4                a
            // 5                b

            List<CT_Row> xrow = sheetData.row;
            Assert.AreEqual(3, xrow.Count);

            // Rows are sorted: [3, 4, 5]
            Assert.AreEqual(3u, xrow[0].r);
            Assert.IsTrue(xrow[0].Equals(row3.GetCTRow()));

            Assert.AreEqual(4u, xrow[1].r);
            Assert.IsTrue(xrow[1].Equals(row1.GetCTRow()));

            Assert.AreEqual(5u, xrow[2].r);
            Assert.IsTrue(xrow[2].Equals(row2.GetCTRow()));
        }

        private void verifyCellContent(ISheet readSheet, int row, String expect)
        {
            IRow readRow = readSheet.GetRow(row);
            if (expect == null)
            {
                Assert.IsNull(readRow);
                return;
            }
            ICell readCell = readRow.GetCell(0);
            if (readCell.CellType == CellType.Numeric)
            {
                Assert.AreEqual(expect, readCell.NumericCellValue.ToString("0.0"));
            }
            else
            {
                Assert.AreEqual(expect, readCell.StringCellValue);
            }
        }
        [Test]
        public void TestBug53798a()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53798.xlsx");

            ISheet testSheet = wb.GetSheetAt(0);
            testSheet.ShiftRows(3, 3, -1);
            foreach (IRow r in testSheet)
            {
                int x = r.RowNum;
            }
            testSheet.ShiftRows(6, 6, 1);

            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            wb.Close();
            Assert.IsNotNull(read);

            ISheet readSheet = read.GetSheetAt(0);
            verifyCellContent(readSheet, 0, "0.0");
            verifyCellContent(readSheet, 1, "1.0");
            verifyCellContent(readSheet, 2, "3.0");
            verifyCellContent(readSheet, 3, null);
            verifyCellContent(readSheet, 4, "4.0");
            verifyCellContent(readSheet, 5, "5.0");
            verifyCellContent(readSheet, 6, null);
            verifyCellContent(readSheet, 7, "6.0");
            verifyCellContent(readSheet, 8, "8.0");
            read.Close();
        }

        [Test]
        public void TestBug56017()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56017.xlsx");

            ISheet sheet = wb.GetSheetAt(0);

            IComment comment = sheet.GetCellComment(new CellAddress(0, 0));
            Assert.IsNotNull(comment);
            Assert.AreEqual("Amdocs", comment.Author);
            Assert.AreEqual("Amdocs:\ntest\n", comment.String.String);

            sheet.ShiftRows(0, 1, 1);

            // comment in row 0 is gone
            comment = sheet.GetCellComment(new CellAddress(0, 0));
            Assert.IsNull(comment);

            // comment is now in row 1
            comment = sheet.GetCellComment(new CellAddress(1, 0));
            Assert.IsNotNull(comment);
            Assert.AreEqual("Amdocs", comment.Author);
            Assert.AreEqual("Amdocs:\ntest\n", comment.String.String);

            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            wb.Close();
            Assert.IsNotNull(wbBack);

            ISheet sheetBack = wbBack.GetSheetAt(0);

            // comment in row 0 is gone
            comment = sheetBack.GetCellComment(new CellAddress(0, 0));
            Assert.IsNull(comment);

            // comment is now in row 1
            comment = sheetBack.GetCellComment(new CellAddress(1, 0));
            Assert.IsNotNull(comment);
            Assert.AreEqual("Amdocs", comment.Author);
            Assert.AreEqual("Amdocs:\ntest\n", comment.String.String);
            wbBack.Close();
        }

        [Test]
        public void Test57171()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            Assert.AreEqual(5, wb.ActiveSheetIndex);
            RemoveAllSheetsBut(5, wb); // 5 is the active / selected sheet
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            IWorkbook wbRead = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            wb.Close();
            Assert.AreEqual(0, wbRead.ActiveSheetIndex);

            wbRead.RemoveSheetAt(0);
            Assert.AreEqual(0, wbRead.ActiveSheetIndex);

            wbRead.Close();
        }

        [Test]
        public void Test57163()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            Assert.AreEqual(5, wb.ActiveSheetIndex);
            wb.RemoveSheetAt(0);
            Assert.AreEqual(4, wb.ActiveSheetIndex);

            wb.Close();
        }

        [Test]
        public void TestSetSheetOrderAndAdjustActiveSheet()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");

            Assert.AreEqual(5, wb.ActiveSheetIndex);

            // Move the sheets around in all possible combinations to check that the active sheet
            // is Set correctly in all cases
            wb.SetSheetOrder(wb.GetSheetName(5), 4);
            Assert.AreEqual(4, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(5), 5);
            Assert.AreEqual(4, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(3), 5);
            Assert.AreEqual(3, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(4), 5);
            Assert.AreEqual(3, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(2), 2);
            Assert.AreEqual(3, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(2), 1);
            Assert.AreEqual(3, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(3), 5);
            Assert.AreEqual(5, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(0), 5);
            Assert.AreEqual(4, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(0), 5);
            Assert.AreEqual(3, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(0), 5);
            Assert.AreEqual(2, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(0), 5);
            Assert.AreEqual(1, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(0), 5);
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.SetSheetOrder(wb.GetSheetName(0), 5);
            Assert.AreEqual(5, wb.ActiveSheetIndex);

            wb.Close();
        }

        [Test]
        public void TestRemoveSheetAndAdjustActiveSheet()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");

            Assert.AreEqual(5, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(0);
            Assert.AreEqual(4, wb.ActiveSheetIndex);

            wb.SetActiveSheet(3);
            Assert.AreEqual(3, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(4);
            Assert.AreEqual(3, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(3);
            Assert.AreEqual(2, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(0);
            Assert.AreEqual(1, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(1);
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(0);
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            try
            {
                wb.RemoveSheetAt(0);
                Assert.Fail("Should catch exception as no more sheets are there");
            }
            catch (ArgumentException)
            {
                // expected
            }
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.CreateSheet();
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(0);
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.Close();
        }

        [Test]
        public void Test57165()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            Assert.AreEqual(5, wb.ActiveSheetIndex);
            RemoveAllSheetsBut(3, wb);
            Assert.AreEqual(0, wb.ActiveSheetIndex);
            wb.CreateSheet("New Sheet1");
            Assert.AreEqual(0, wb.ActiveSheetIndex);
            wb.CloneSheet(0); // Throws exception here
            wb.SetSheetName(1, "New Sheet");
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.Close();
        }

        private static void RemoveAllSheetsBut(int sheetIndex, IWorkbook wb)
        {
            int sheetNb = wb.NumberOfSheets;
            // Move this sheet at the first position
            wb.SetSheetOrder(wb.GetSheetName(sheetIndex), 0);
            // Must make this sheet active (otherwise, for XLSX, Excel might protest that active sheet no longer exists)
            // I think POI should automatically handle this case when deleting sheets...
            //      wb.SetActiveSheet(0);
            for (int sn = sheetNb - 1; sn > 0; sn--)
            {
                wb.RemoveSheetAt(sn);
            }
        }

        [Test]
        public void TestBug57828_OnlyOneCommentShiftedInRow()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57828.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            IComment comment1 = sheet.GetCellComment(new CellAddress(2, 1));
            Assert.IsNotNull(comment1);

            IComment comment2 = sheet.GetCellComment(new CellAddress(2, 2));
            Assert.IsNotNull(comment2);

            IComment comment3 = sheet.GetCellComment(new CellAddress(1, 1));
            Assert.IsNull(comment3, "NO comment in (1,1) and it should be null");

            sheet.ShiftRows(2, 2, -1);

            comment3 = sheet.GetCellComment(new CellAddress(1, 1));
            Assert.IsNotNull(comment3, "Comment in (2,1) Moved to (1,1) so its not null now.");

            comment1 = sheet.GetCellComment(new CellAddress(2, 1));
            Assert.IsNull(comment1, "No comment currently in (2,1) and hence it is null");

            comment2 = sheet.GetCellComment(new CellAddress(1, 2));
            Assert.IsNotNull(comment2, "Comment in (2,2) should have Moved as well because of shift rows. But its not");

            wb.Close();
        }

        private static String getCellFormula(ISheet sheet, String address)
        {
            CellAddress cellAddress = new CellAddress(address);
            IRow row = sheet.GetRow(cellAddress.Row);
            Assert.IsNotNull(row);
            ICell cell = row.GetCell(cellAddress.Column);
            Assert.IsNotNull(cell);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            return cell.CellFormula;
        }
        // This test is written as expected-to-Assert.Fail and should be rewritten
        // as expected-to-pass when the bug is fixed.
        [Test]
        public void TestSharedFormulas()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("TestShiftRowSharedFormula.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            Assert.AreEqual("SUM(C2:C4)", getCellFormula(sheet, "C5"));
            Assert.AreEqual("SUM(D2:D4)", getCellFormula(sheet, "D5"));
            Assert.AreEqual("SUM(E2:E4)", getCellFormula(sheet, "E5"));
            sheet.ShiftRows(3, sheet.LastRowNum, 1);
            try
            {
                Assert.AreEqual("SUM(C2:C5)", getCellFormula(sheet, "C6"));
                Assert.AreEqual("SUM(D2:D5)", getCellFormula(sheet, "D6"));
                Assert.AreEqual("SUM(E2:E5)", getCellFormula(sheet, "E6"));
                POITestCase.TestPassesNow(59983);
            }
            catch (AssertionException e)
            {
                POITestCase.SkipTest(e);
            }

            wb.Close();
        }

        // bug 60260: shift rows or rename a sheet containing a named range
        // that refers to formula with a unicode (non-ASCII) sheet name formula
        [Test]
        public void ShiftRowsWithUnicodeNamedRange()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("unicodeSheetName.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            sheet.ShiftRows(1, 2, 3);
            IOUtils.CloseQuietly(wb);
        }
    }
}

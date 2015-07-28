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

using System;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NUnit.Framework;
using TestCases.SS.UserModel;
namespace NPOI.XSSF.UserModel
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
        public override void TestShiftRowBreaks() { // disabled test from superclass
            // TODO - support shifting of page breaks
        }

        [Test]
        public void TestBug54524()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("54524.xlsx");
            ISheet sheet = workbook.GetSheetAt(0);
            sheet.ShiftRows(3, 5, -1);

            ICell cell = CellUtil.GetCell(sheet.GetRow(1), 0);
            Assert.AreEqual(1.0, cell.NumericCellValue);
            cell = CellUtil.GetCell(sheet.GetRow(2), 0);
            Assert.AreEqual("SUM(A2:A2)", cell.CellFormula);
            cell = CellUtil.GetCell(sheet.GetRow(3), 0);
            Assert.AreEqual("X", cell.StringCellValue);
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
            testSheet.ShiftRows(6, 7, 1);	// -- CHANGE the shift to positive once the behaviour of  
            // the above has been tested

            //saveReport(wb, new File("/tmp/53798.xlsx"));
            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(wb);
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

            //saveReport(wb, new File("/tmp/53798.xlsx"));
            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(wb);
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
        }

        [Test]
        public void TestBug56017()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56017.xlsx");

            ISheet sheet = wb.GetSheetAt(0);

            IComment comment = sheet.GetCellComment(0, 0);
            Assert.IsNotNull(comment);
            Assert.AreEqual("Amdocs", comment.Author);
            Assert.AreEqual("Amdocs:\ntest\n", comment.String.String);

            sheet.ShiftRows(0, 1, 1);

            // comment in row 0 is gone
            comment = sheet.GetCellComment(0, 0);
            Assert.IsNull(comment);

            // comment is now in row 1
            comment = sheet.GetCellComment(1, 0);
            Assert.IsNotNull(comment);
            Assert.AreEqual("Amdocs", comment.Author);
            Assert.AreEqual("Amdocs:\ntest\n", comment.String.String);

            //        FileOutputStream outputStream = new FileOutputStream("/tmp/56017.xlsx");
            //        try {
            //            wb.Write(outputStream);
            //        } finally {
            //            outputStream.Close();
            //        }

            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);

            ISheet sheetBack = wbBack.GetSheetAt(0);

            // comment in row 0 is gone
            comment = sheetBack.GetCellComment(0, 0);
            Assert.IsNull(comment);

            // comment is now in row 1
            comment = sheetBack.GetCellComment(1, 0);
            Assert.IsNotNull(comment);
            Assert.AreEqual("Amdocs", comment.Author);
            Assert.AreEqual("Amdocs:\ntest\n", comment.String.String);
        }

        [Test]
        public void Test57171()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            Assert.AreEqual(5, wb.ActiveSheetIndex);
            RemoveAllSheetsBut(5, wb); // 5 is the active / selected sheet
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            IWorkbook wbRead = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.AreEqual(0, wbRead.ActiveSheetIndex);

            wbRead.RemoveSheetAt(0);
            Assert.AreEqual(0, wbRead.ActiveSheetIndex);

            //wb.Write(new FileOutputStream("/tmp/57171.xls"));
        }

        [Test]
        public void Test57163()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            Assert.AreEqual(5, wb.ActiveSheetIndex);
            wb.RemoveSheetAt(0);
            Assert.AreEqual(4, wb.ActiveSheetIndex);

            //wb.Write(new FileOutputStream("/tmp/57163.xls"));
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
            catch (ArgumentException e)
            {
                // expected
            }
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.CreateSheet();
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            wb.RemoveSheetAt(0);
            Assert.AreEqual(0, wb.ActiveSheetIndex);
        }

        // TODO: enable when bug 57165 is fixed
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

            //wb.Write(new FileOutputStream("/tmp/57165.xls"));
        }

        //    public void Test57165b(){
        //        IWorkbook wb = new XSSFWorkbook();
        //        try {
        //            wb.CreateSheet("New Sheet 1");
        //            wb.CreateSheet("New Sheet 2");
        //        } finally {
        //            wb.Close();
        //        }
        //    }

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

            IComment comment1 = sheet.GetCellComment(2, 1);
            Assert.IsNotNull(comment1);

            IComment comment2 = sheet.GetCellComment(2, 2);
            Assert.IsNotNull(comment2);

            IComment comment3 = sheet.GetCellComment(1, 1);
            Assert.IsNull(comment3, "NO comment in (1,1) and it should be null");

            sheet.ShiftRows(2, 2, -1);

            comment3 = sheet.GetCellComment(1, 1);
            Assert.IsNotNull(comment3, "Comment in (2,1) Moved to (1,1) so its not null now.");

            comment1 = sheet.GetCellComment(2, 1);
            Assert.IsNull(comment1, "No comment currently in (2,1) and hence it is null");

            comment2 = sheet.GetCellComment(1, 2);
            Assert.IsNotNull(comment2, "Comment in (2,2) should have Moved as well because of shift rows. But its not");

            //        OutputStream stream = new FileOutputStream("/tmp/57828.xlsx");
            //        try {
            //        	wb.Write(stream);
            //        } finally {
            //        	stream.Close();
            //        }

            wb.Close();
        }

    }
}


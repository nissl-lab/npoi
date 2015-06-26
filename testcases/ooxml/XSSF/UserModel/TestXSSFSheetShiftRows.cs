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

using TestCases.SS.UserModel;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
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
    }
}


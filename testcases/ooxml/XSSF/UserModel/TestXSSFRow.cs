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

using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using TestCases.SS.UserModel;

namespace TestCases.XSSF.UserModel
{

    /**
     * Tests for XSSFRow
     */
    [TestFixture]
    public class TestXSSFRow : BaseTestXRow
    {

        public TestXSSFRow():base(XSSFITestDataProvider.instance)
        {
            
        }

        [Test]
        public void TestCopyRowFrom()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.CreateSheet("test") as XSSFSheet;
            XSSFRow srcRow = sheet.CreateRow(0) as XSSFRow;
            srcRow.CreateCell(0).SetCellValue("Hello");
            XSSFRow destRow = sheet.CreateRow(1) as XSSFRow;

            destRow.CopyRowFrom(srcRow, new CellCopyPolicy());
            Assert.IsNotNull(destRow.GetCell(0));
            Assert.AreEqual("Hello", destRow.GetCell(0).StringCellValue);

            workbook.Close();
        }
        [Test]
        public void TestCopyRowFromExternalSheet()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet srcSheet = workbook.CreateSheet("src");
            XSSFSheet destSheet = workbook.CreateSheet("dest") as XSSFSheet;
            workbook.CreateSheet("other");

            IRow srcRow = srcSheet.CreateRow(0);
            int col = 0;
            //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
            srcRow.CreateCell(col++).CellFormula = ("B5");
            srcRow.CreateCell(col++).CellFormula = ("src!B5");
            srcRow.CreateCell(col++).CellFormula = ("dest!B5");
            srcRow.CreateCell(col++).CellFormula = ("other!B5");

            //Test 2D and 3D Ref Ptgs with absolute row
            srcRow.CreateCell(col++).CellFormula = ("B$5");
            srcRow.CreateCell(col++).CellFormula = ("src!B$5");
            srcRow.CreateCell(col++).CellFormula = ("dest!B$5");
            srcRow.CreateCell(col++).CellFormula = ("other!B$5");

            //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
            srcRow.CreateCell(col++).CellFormula = ("SUM(B5:D$5)");
            srcRow.CreateCell(col++).CellFormula = ("SUM(src!B5:D$5)");
            srcRow.CreateCell(col++).CellFormula = ("SUM(dest!B5:D$5)");
            srcRow.CreateCell(col++).CellFormula = ("SUM(other!B5:D$5)");
            //////////////////
            XSSFRow destRow = destSheet.CreateRow(1) as XSSFRow;
            destRow.CopyRowFrom(srcRow, new CellCopyPolicy());

            //////////////////

            //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
            col = 0;
            ICell cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("B6", cell.CellFormula, "RefPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("src!B6", cell.CellFormula, "Ref3DPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("dest!B6", cell.CellFormula, "Ref3DPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("other!B6", cell.CellFormula, "Ref3DPtg");

            /////////////////////////////////////////////

            //Test 2D and 3D Ref Ptgs with absolute row (Ptg row number shouldn't change)
            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("B$5", cell.CellFormula, "RefPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("src!B$5", cell.CellFormula, "Ref3DPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("dest!B$5", cell.CellFormula, "Ref3DPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("other!B$5", cell.CellFormula, "Ref3DPtg") ;

            //////////////////////////////////////////

            //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
            // Note: absolute row changes from last cell to first cell in order
            // to maintain topLeft:bottomRight order
            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("SUM(B$5:D6)", cell.CellFormula, "Area2DPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("SUM(src!B$5:D6)", cell.CellFormula, "Area3DPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(destRow.GetCell(6));
            Assert.AreEqual( "SUM(dest!B$5:D6)", cell.CellFormula, "Area3DPtg");

            cell = destRow.GetCell(col++);
            Assert.IsNotNull(destRow.GetCell(7));
            Assert.AreEqual( "SUM(other!B$5:D6)", cell.CellFormula, "Area3DPtg");

            workbook.Close();
        }
        [Test]
        public void TestCopyRowOverwritesExistingRow()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = workbook.CreateSheet("Sheet1") as XSSFSheet;
            ISheet sheet2 = workbook.CreateSheet("Sheet2");

            IRow srcRow = sheet1.CreateRow(0);
            XSSFRow destRow = sheet1.CreateRow(1) as XSSFRow;
            IRow observerRow = sheet1.CreateRow(2);
            IRow externObserverRow = sheet2.CreateRow(0);

            srcRow.CreateCell(0).SetCellValue("hello");
            srcRow.CreateCell(1).SetCellValue("world");
            destRow.CreateCell(0).SetCellValue(5.0); //A2 -> 5.0
            destRow.CreateCell(1).CellFormula = ("A1"); // B2 -> A1 -> "hello"
            observerRow.CreateCell(0).CellFormula = ("A2"); // A3 -> A2 -> 5.0
            observerRow.CreateCell(1).CellFormula = ("B2"); // B3 -> B2 -> A1 -> "hello"
            externObserverRow.CreateCell(0).CellFormula = ("Sheet1!A2"); //Sheet2!A1 -> Sheet1!A2 -> 5.0

            // overwrite existing destRow with row-copy of srcRow
            destRow.CopyRowFrom(srcRow, new CellCopyPolicy());

            // copyRowFrom should update existing destRow, rather than creating a new row and reassigning the destRow pointer
            // to the new row (and allow the old row to be garbage collected)
            // this is mostly so existing references to rows that are overwritten are updated
            // rather than allowing users to continue updating rows that are no longer part of the sheet

            Assert.AreSame(srcRow, sheet1.GetRow(0), "existing references to srcRow are still valid");
            Assert.AreSame(destRow, sheet1.GetRow(1), "existing references to destRow are still valid");
            Assert.AreSame(observerRow, sheet1.GetRow(2), "existing references to observerRow are still valid");
            Assert.AreSame(externObserverRow, sheet2.GetRow(0), "existing references to externObserverRow are still valid");

            // Make sure copyRowFrom actually copied row (this is tested elsewhere)
            Assert.AreEqual(CellType.String, destRow.GetCell(0).CellType);
            Assert.AreEqual("hello", destRow.GetCell(0).StringCellValue);

            // We don't want #REF! errors if we copy a row that contains cells that are referred to by other cells outside of copied region
            Assert.AreEqual("A2", observerRow.GetCell(0).CellFormula, "references to overwritten cells are unmodified");
            Assert.AreEqual("B2", observerRow.GetCell(1).CellFormula, "references to overwritten cells are unmodified");
            Assert.AreEqual("Sheet1!A2", externObserverRow.GetCell(0).CellFormula, "references to overwritten cells are unmodified");

            workbook.Close();
        }

    }

}
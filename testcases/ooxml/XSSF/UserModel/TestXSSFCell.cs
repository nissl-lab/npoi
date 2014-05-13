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

using NUnit.Framework;
using TestCases.SS.UserModel;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using NPOI.XSSF.Model;
namespace NPOI.XSSF.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFCell : BaseTestCell
    {

        public TestXSSFCell()
            : base(XSSFITestDataProvider.instance)
        {

        }

        /**
         * Bug 47026: trouble changing cell type when workbook doesn't contain
         * Shared String Table
         */
        [Test]
        public void Test47026_1()
        {
            IWorkbook source = _testDataProvider.OpenSampleWorkbook("47026.xlsm");
            ISheet sheet = source.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            cell.SetCellType(CellType.String);
            cell.SetCellValue("456");
        }
        [Test]
        public void Test47026_2()
        {
            IWorkbook source = _testDataProvider.OpenSampleWorkbook("47026.xlsm");
            ISheet sheet = source.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            cell.SetCellFormula(null);
            cell.SetCellValue("456");
        }

        /**
         * Test that we can read inline strings that are expressed directly in the cell defInition
         * instead of implementing the shared string table.
         *
         * Some programs, for example, Microsoft Excel Driver for .xlsx insert inline string
         * instead of using the shared string table. See bug 47206
         */
        [Test]
        public void TestInlineString()
        {
            XSSFWorkbook wb = (XSSFWorkbook)_testDataProvider.OpenSampleWorkbook("xlsx-jdbc.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            XSSFRow row = (XSSFRow)sheet.GetRow(1);

            XSSFCell cell_0 = (XSSFCell)row.GetCell(0);
            Assert.AreEqual(ST_CellType.inlineStr, cell_0.GetCTCell().t);
            Assert.IsTrue(cell_0.GetCTCell().IsSetIs());
            Assert.AreEqual(cell_0.StringCellValue, "A Very large string in column 1 AAAAAAAAAAAAAAAAAAAAA");

            XSSFCell cell_1 = (XSSFCell)row.GetCell(1);
            Assert.AreEqual(ST_CellType.inlineStr, cell_1.GetCTCell().t);
            Assert.IsTrue(cell_1.GetCTCell().IsSetIs());
            Assert.AreEqual(cell_1.StringCellValue, "foo");

            XSSFCell cell_2 = (XSSFCell)row.GetCell(2);
            Assert.AreEqual(ST_CellType.inlineStr, cell_2.GetCTCell().t);
            Assert.IsTrue(cell_2.GetCTCell().IsSetIs());
            Assert.AreEqual(row.GetCell(2).StringCellValue, "bar");
        }

        /**
         *  Bug 47278 -  xsi:nil attribute for <t> tag caused Excel 2007 to fail to open workbook
         */
        [Test]
        public void Test47278()
        {
            XSSFWorkbook wb = (XSSFWorkbook)_testDataProvider.CreateWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            SharedStringsTable sst = wb.GetSharedStringSource();
            Assert.AreEqual(0, sst.Count);

            //case 1. cell.SetCellValue(new XSSFRichTextString((String)null));
            ICell cell_0 = row.CreateCell(0);
            XSSFRichTextString str = new XSSFRichTextString((String)null);
            Assert.IsNull(str.String);
            cell_0.SetCellValue(str);
            Assert.AreEqual(0, sst.Count);
            Assert.AreEqual(CellType.Blank, cell_0.CellType);

            //case 2. cell.SetCellValue((String)null);
            ICell cell_1 = row.CreateCell(1);
            cell_1.SetCellValue((String)null);
            Assert.AreEqual(0, sst.Count);
            Assert.AreEqual(CellType.Blank, cell_1.CellType);
        }
        [Test]
        public void TestFormulaString()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFCell cell = (XSSFCell)wb.CreateSheet().CreateRow(0).CreateCell(0);
            CT_Cell ctCell = cell.GetCTCell(); //low-level bean holding cell's xml

            cell.SetCellFormula("A2");
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(cell.CellFormula, "A2");
            //the value is not Set and cell's type='N' which means blank
            Assert.AreEqual(ST_CellType.n, ctCell.t);

            //set cached formula value
            cell.SetCellValue("t='str'");
            //we are still of 'formula' type
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(cell.CellFormula, "A2");
            //cached formula value is Set and cell's type='STR'
            Assert.AreEqual(ST_CellType.str, ctCell.t);
            Assert.AreEqual(cell.StringCellValue, "t='str'");

            //now remove the formula, the cached formula result remains
            cell.SetCellFormula(null);
            Assert.AreEqual(CellType.String, cell.CellType);
            Assert.AreEqual(ST_CellType.str, ctCell.t);
            //the line below failed prior to fix of Bug #47889
            Assert.AreEqual(cell.StringCellValue, "t='str'");

            //revert to a blank cell
            cell.SetCellValue((String)null);
            Assert.AreEqual(CellType.Blank, cell.CellType);
            Assert.AreEqual(ST_CellType.n, ctCell.t);
            Assert.AreEqual(cell.StringCellValue, "");
        }

        /**
         * Bug 47889: problems when calling XSSFCell.StringCellValue on a workbook Created in Gnumeric
         */
        [Test]
        public void Test47889()
        {
            XSSFWorkbook wb = (XSSFWorkbook)_testDataProvider.OpenSampleWorkbook("47889.xlsx");
            ISheet sh = wb.GetSheetAt(0);

            ICell cell;

            //try a string cell
            cell = sh.GetRow(0).GetCell(0);
            Assert.AreEqual(CellType.String, cell.CellType);
            Assert.AreEqual(cell.StringCellValue, "a");
            Assert.AreEqual(cell.ToString(), "a");
            //Gnumeric produces spreadsheets without styles
            //make sure we return null for that instead of throwing OutOfBounds
            Assert.AreEqual(null, cell.CellStyle);

            //try a numeric cell
            cell = sh.GetRow(1).GetCell(0);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            Assert.AreEqual(1.0, cell.NumericCellValue);
            Assert.AreEqual(cell.ToString(), "1");
            //Gnumeric produces spreadsheets without styles
            //make sure we return null for that instead of throwing OutOfBounds
            Assert.AreEqual(null, cell.CellStyle);
        }

        /**
         * Cell with the formula that returns error must return error code(There was
         * an problem that cell could not return error value form formula cell).
         */
        [Test]
        public void TestGetErrorCellValueFromFormulaCell()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellFormula("SQRT(-1)");
            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(cell);
            Assert.AreEqual(36, cell.ErrorCellValue);
        }
        [Test]
        public void TestIsMergedCell()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(5);
            ICell cell5 = row.CreateCell(5);
            ICell cell6 = row.CreateCell(6);
            ICell cell8 = row.CreateCell(8);

            Assert.IsFalse(cell5.IsMergedCell);
            Assert.IsFalse(cell6.IsMergedCell);
            Assert.IsFalse(cell8.IsMergedCell);

            sheet.AddMergedRegion(new SS.Util.CellRangeAddress(5, 6, 5, 6));   //region with 4 cells

            Assert.IsTrue(cell5.IsMergedCell);
            Assert.IsTrue(cell6.IsMergedCell);
            Assert.IsFalse(cell8.IsMergedCell);

        }
        [Test]
        public void TestMissingRAttribute()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFRow row = (XSSFRow)sheet.CreateRow(0);
            XSSFCell a1 = (XSSFCell)row.CreateCell(0);
            a1.SetCellValue("A1");
            XSSFCell a2 = (XSSFCell)row.CreateCell(1);
            a2.SetCellValue("B1");
            XSSFCell a4 = (XSSFCell)row.CreateCell(4);
            a4.SetCellValue("E1");
            XSSFCell a6 = (XSSFCell)row.CreateCell(5);
            a6.SetCellValue("F1");

            assertCellsWithMissingR(row);

            a2.GetCTCell().unsetR();
            a6.GetCTCell().unsetR();

            assertCellsWithMissingR(row);

            wb = (XSSFWorkbook)_testDataProvider.WriteOutAndReadBack(wb);
            row = (XSSFRow)wb.GetSheetAt(0).GetRow(0);
            assertCellsWithMissingR(row);
        }

        private void assertCellsWithMissingR(XSSFRow row)
        {
            XSSFCell a1 = (XSSFCell)row.GetCell(0);
            Assert.IsNotNull(a1);
            XSSFCell a2 = (XSSFCell)row.GetCell(1);
            Assert.IsNotNull(a2);
            XSSFCell a5 = (XSSFCell)row.GetCell(4);
            Assert.IsNotNull(a5);
            XSSFCell a6 = (XSSFCell)row.GetCell(5);
            Assert.IsNotNull(a6);

            Assert.AreEqual(6, row.LastCellNum);
            Assert.AreEqual(4, row.PhysicalNumberOfCells);

            Assert.AreEqual(a1.StringCellValue, "A1");
            Assert.AreEqual(a2.StringCellValue, "B1");
            Assert.AreEqual(a5.StringCellValue, "E1");
            Assert.AreEqual(a6.StringCellValue, "F1");

            // even if R attribute is not set,
            // POI is able to re-construct it from column and row indexes
            Assert.AreEqual(a1.GetReference(), "A1");
            Assert.AreEqual(a2.GetReference(), "B1");
            Assert.AreEqual(a5.GetReference(), "E1");
            Assert.AreEqual(a6.GetReference(), "F1");
        }
        [Test]
        public void TestMissingRAttributeBug54288()
        {
            // workbook with cells missing the R attribute
            XSSFWorkbook wb = (XSSFWorkbook)_testDataProvider.OpenSampleWorkbook("54288.xlsx");
            // same workbook re-saved in Excel 2010, the R attribute is updated for every cell with the right value.
            XSSFWorkbook wbRef = (XSSFWorkbook)_testDataProvider.OpenSampleWorkbook("54288-ref.xlsx");

            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            XSSFSheet sheetRef = (XSSFSheet)wbRef.GetSheetAt(0);
            Assert.AreEqual(sheetRef.PhysicalNumberOfRows, sheet.PhysicalNumberOfRows);

            // Test idea: iterate over cells in the reference worksheet, they all have the R attribute set.
            // For each cell from the reference sheet find the corresponding cell in the problematic file (with missing R)
            // and assert that POI reads them equally:
            DataFormatter formater = new DataFormatter();
            foreach (IRow r in sheetRef)
            {
                XSSFRow rowRef = (XSSFRow)r;
                XSSFRow row = (XSSFRow)sheet.GetRow(rowRef.RowNum);

                Assert.AreEqual(rowRef.PhysicalNumberOfCells, row.PhysicalNumberOfCells, "number of cells in row[" + row.RowNum + "]");

                foreach (ICell c in rowRef.Cells)
                {
                    XSSFCell cellRef = (XSSFCell)c;
                    XSSFCell cell = (XSSFCell)row.GetCell(cellRef.ColumnIndex);

                    Assert.AreEqual(cellRef.ColumnIndex, cell.ColumnIndex);
                    Assert.AreEqual(cellRef.GetReference(), cell.GetReference());

                    if (!cell.GetCTCell().IsSetR())
                    {
                        Assert.IsTrue(cellRef.GetCTCell().IsSetR(), "R must e set in cellRef");

                        String valRef = formater.FormatCellValue(cellRef);
                        String val = formater.FormatCellValue(cell);
                        Assert.AreEqual(valRef, val);
                    }

                }
            }
        }
    }

}
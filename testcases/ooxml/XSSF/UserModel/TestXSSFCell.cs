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
            cell.SetCellType(CellType.STRING);
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
            Assert.AreEqual("A Very large string in column 1 AAAAAAAAAAAAAAAAAAAAA", cell_0.StringCellValue);

            XSSFCell cell_1 = (XSSFCell)row.GetCell(1);
            Assert.AreEqual(ST_CellType.inlineStr, cell_1.GetCTCell().t);
            Assert.IsTrue(cell_1.GetCTCell().IsSetIs());
            Assert.AreEqual("foo", cell_1.StringCellValue);

            XSSFCell cell_2 = (XSSFCell)row.GetCell(2);
            Assert.AreEqual(ST_CellType.inlineStr, cell_2.GetCTCell().t);
            Assert.IsTrue(cell_2.GetCTCell().IsSetIs());
            Assert.AreEqual("bar", row.GetCell(2).StringCellValue);
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
            Assert.AreEqual(0, sst.GetCount());

            //case 1. cell.SetCellValue(new XSSFRichTextString((String)null));
            ICell cell_0 = row.CreateCell(0);
            XSSFRichTextString str = new XSSFRichTextString((String)null);
            Assert.IsNull(str.String);
            cell_0.SetCellValue(str);
            Assert.AreEqual(0, sst.GetCount());
            Assert.AreEqual(CellType.BLANK, cell_0.CellType);

            //case 2. cell.SetCellValue((String)null);
            ICell cell_1 = row.CreateCell(1);
            cell_1.SetCellValue((String)null);
            Assert.AreEqual(0, sst.GetCount());
            Assert.AreEqual(CellType.BLANK, cell_1.CellType);
        }
        [Test]
        public void TestFormulaString()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFCell cell = (XSSFCell)wb.CreateSheet().CreateRow(0).CreateCell(0);
            CT_Cell ctCell = cell.GetCTCell(); //low-level bean holding cell's xml

            cell.SetCellFormula("A2");
            Assert.AreEqual(CellType.FORMULA, cell.CellType);
            Assert.AreEqual("A2", cell.CellFormula);
            //the value is not Set and cell's type='N' which means blank
            Assert.AreEqual(ST_CellType.n, ctCell.t);

            //set cached formula value
            cell.SetCellValue("t='str'");
            //we are still of 'formula' type
            Assert.AreEqual(CellType.FORMULA, cell.CellType);
            Assert.AreEqual("A2", cell.CellFormula);
            //cached formula value is Set and cell's type='STR'
            Assert.AreEqual(ST_CellType.str, ctCell.t);
            Assert.AreEqual("t='str'", cell.StringCellValue);

            //now remove the formula, the cached formula result remains
            cell.SetCellFormula(null);
            Assert.AreEqual(CellType.STRING, cell.CellType);
            Assert.AreEqual(ST_CellType.str, ctCell.t);
            //the line below failed prior to fix of Bug #47889
            Assert.AreEqual("t='str'", cell.StringCellValue);

            //revert to a blank cell
            cell.SetCellValue((String)null);
            Assert.AreEqual(CellType.BLANK, cell.CellType);
            Assert.AreEqual(ST_CellType.n, ctCell.t);
            Assert.AreEqual("", cell.StringCellValue);
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
            Assert.AreEqual(CellType.STRING, cell.CellType);
            Assert.AreEqual("a", cell.StringCellValue);
            Assert.AreEqual("a", cell.ToString());
            //Gnumeric produces spreadsheets without styles
            //make sure we return null for that instead of throwing OutOfBounds
            Assert.AreEqual(null, cell.CellStyle);

            //try a numeric cell
            cell = sh.GetRow(1).GetCell(0);
            Assert.AreEqual(CellType.NUMERIC, cell.CellType);
            Assert.AreEqual(1.0, cell.NumericCellValue);
            Assert.AreEqual("1", cell.ToString());
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
    }

}
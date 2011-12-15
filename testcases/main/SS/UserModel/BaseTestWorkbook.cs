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

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    /**
     * @author Yegor Kozlov
     */
    [TestClass]
    public class BaseTestWorkbook
    {

       private ITestDataProvider _testDataProvider;

        public BaseTestWorkbook()
       {
           _testDataProvider = TestCases.HSSF.HSSFITestDataProvider.Instance;
       }

       protected BaseTestWorkbook(ITestDataProvider testDataProvider)
       {
             _testDataProvider = testDataProvider;
        }

        [TestMethod]
        public void TestCreateSheet()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, wb.NumberOfSheets);

            //Getting a sheet by invalid index or non-existing name
            Assert.IsNull(wb.GetSheet("Sheet1"));
            try
            {
                wb.GetSheetAt(0);
                Assert.Fail("should have thrown exceptiuon due to invalid sheet index");
            }
            catch (ArgumentException)
            {
                // expected during successful test
            }

            ISheet sheet0 = wb.CreateSheet();
            ISheet sheet1 = wb.CreateSheet();
            Assert.AreEqual("Sheet0", sheet0.SheetName);
            Assert.AreEqual("Sheet1", sheet1.SheetName);
            Assert.AreEqual(2, wb.NumberOfSheets);

            //fetching sheets by name is case-insensitive
            ISheet originalSheet = wb.CreateSheet("Sheet3");
            ISheet fetchedSheet = wb.GetSheet("sheet3");
            if (fetchedSheet == null)
            {
                throw new AssertFailedException("Identified bug 44892");
            }
            Assert.AreEqual("Sheet3", fetchedSheet.SheetName);
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.AreSame(originalSheet, fetchedSheet);
            try
            {
                wb.CreateSheet("sHeeT3");
                Assert.Fail("should have thrown exceptiuon due to duplicate sheet name");
            }
            catch (ArgumentException e)
            {
                // expected during successful test
                Assert.AreEqual("The workbook already contains a sheet of this name", e.Message);
            }

            //names cannot be blank or contain any of /\*?[]
            String[] invalidNames = {"", "Sheet/", "Sheet\\",
                "Sheet?", "Sheet*", "Sheet[", "Sheet]"};
            foreach (String sheetName in invalidNames)
            {
                try
                {
                    wb.CreateSheet(sheetName);
                    Assert.Fail("should have thrown exception due to invalid sheet name: " + sheetName);
                }
                catch (ArgumentException)
                {
                    // expected during successful test
                }
            }
            //still have 3 sheets
            Assert.AreEqual(3, wb.NumberOfSheets);

            //change the name of the 3rd sheet
            wb.SetSheetName(2, "I Changed!");

            //try to assign an invalid name to the 2nd sheet
            try
            {
                wb.SetSheetName(1, "[I'm invalid]");
                Assert.Fail("should have thrown exceptiuon due to invalid sheet name");
            }
            catch (ArgumentException)
            {
                // expected during successful test
            }

            //check
            Assert.AreEqual(0, wb.GetSheetIndex("sheet0"));
            Assert.AreEqual(1, wb.GetSheetIndex("sheet1"));
            Assert.AreEqual(2, wb.GetSheetIndex("I Changed!"));

            Assert.AreSame(sheet0, wb.GetSheet("sheet0"));
            Assert.AreSame(sheet1, wb.GetSheet("sheet1"));
            Assert.AreSame(originalSheet, wb.GetSheet("I Changed!"));
            Assert.IsNull(wb.GetSheet("unknown"));

            //serialize and read again
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.AreEqual(0, wb.GetSheetIndex("sheet0"));
            Assert.AreEqual(1, wb.GetSheetIndex("sheet1"));
            Assert.AreEqual(2, wb.GetSheetIndex("I Changed!"));
        }
        [TestMethod]
        public void TestRemoveSheetAt()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("sheet1");
            workbook.CreateSheet("sheet2");
            workbook.CreateSheet("sheet3");
            Assert.AreEqual(3, workbook.NumberOfSheets);
            workbook.RemoveSheetAt(1);
            Assert.AreEqual(2, workbook.NumberOfSheets);
            Assert.AreEqual("sheet3", workbook.GetSheetName(1));
            workbook.RemoveSheetAt(0);
            Assert.AreEqual(1, workbook.NumberOfSheets);
            Assert.AreEqual("sheet3", workbook.GetSheetName(0));
            workbook.RemoveSheetAt(0);
            Assert.AreEqual(0, workbook.NumberOfSheets);

            //re-create the sheets
            workbook.CreateSheet("sheet1");
            workbook.CreateSheet("sheet2");
            workbook.CreateSheet("sheet3");
            Assert.AreEqual(3, workbook.NumberOfSheets);
        }
        [TestMethod]
        public void TestDefaultValues()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, b.ActiveSheetIndex);
            Assert.AreEqual(0, b.FirstVisibleTab);
            Assert.AreEqual(0, b.NumberOfNames);
            Assert.AreEqual(0, b.NumberOfSheets);
        }
        [TestMethod]
        public void TestSheetSelection()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            b.CreateSheet("Sheet One");
            b.CreateSheet("Sheet Two");
            b.SetActiveSheet(1);
            b.SetSelectedTab(1);
            b.FirstVisibleTab = (1);
            Assert.AreEqual(1, b.ActiveSheetIndex);
            Assert.AreEqual(1, b.FirstVisibleTab);
        }
        [TestMethod]
        public void TestPrintArea()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Test Print Area");
            String sheetName1 = sheet1.SheetName;

            // workbook.SetPrintArea(0, reference);
            workbook.SetPrintArea(0, 1, 5, 4, 9);
            String retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.AreEqual("'" + sheetName1 + "'!$B$5:$F$10", retrievedPrintArea);

            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);
            retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.AreEqual("'" + sheetName1 + "'!" + reference, retrievedPrintArea);

            workbook.RemovePrintArea(0);
            Assert.IsNull(workbook.GetPrintArea(0));
        }
        [TestMethod]
        public void TestGetSetActiveSheet()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, workbook.ActiveSheetIndex);

            workbook.CreateSheet("sheet1");
            workbook.CreateSheet("sheet2");
            workbook.CreateSheet("sheet3");
            // set second sheet
            workbook.SetActiveSheet(1);
            // test if second sheet is set up
            Assert.AreEqual(1, workbook.ActiveSheetIndex);

            workbook.SetActiveSheet(0);
            // test if second sheet is set up
            Assert.AreEqual(0, workbook.ActiveSheetIndex);
        }
        [TestMethod]
        public void TestSetSheetOrder()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            for (int i = 0; i < 10; i++)
            {
                wb.CreateSheet("Sheet " + i);
            }

            // Check the Initial order
            Assert.AreEqual(0, wb.GetSheetIndex("Sheet 0"));
            Assert.AreEqual(1, wb.GetSheetIndex("Sheet 1"));
            Assert.AreEqual(2, wb.GetSheetIndex("Sheet 2"));
            Assert.AreEqual(3, wb.GetSheetIndex("Sheet 3"));
            Assert.AreEqual(4, wb.GetSheetIndex("Sheet 4"));
            Assert.AreEqual(5, wb.GetSheetIndex("Sheet 5"));
            Assert.AreEqual(6, wb.GetSheetIndex("Sheet 6"));
            Assert.AreEqual(7, wb.GetSheetIndex("Sheet 7"));
            Assert.AreEqual(8, wb.GetSheetIndex("Sheet 8"));
            Assert.AreEqual(9, wb.GetSheetIndex("Sheet 9"));

            // Change
            wb.SetSheetOrder("Sheet 6", 0);
            wb.SetSheetOrder("Sheet 3", 7);
            wb.SetSheetOrder("Sheet 1", 9);

            // Check they're currently right
            Assert.AreEqual(0, wb.GetSheetIndex("Sheet 6"));
            Assert.AreEqual(1, wb.GetSheetIndex("Sheet 0"));
            Assert.AreEqual(2, wb.GetSheetIndex("Sheet 2"));
            Assert.AreEqual(3, wb.GetSheetIndex("Sheet 4"));
            Assert.AreEqual(4, wb.GetSheetIndex("Sheet 5"));
            Assert.AreEqual(5, wb.GetSheetIndex("Sheet 7"));
            Assert.AreEqual(6, wb.GetSheetIndex("Sheet 3"));
            Assert.AreEqual(7, wb.GetSheetIndex("Sheet 8"));
            Assert.AreEqual(8, wb.GetSheetIndex("Sheet 9"));
            Assert.AreEqual(9, wb.GetSheetIndex("Sheet 1"));

            IWorkbook wbr = _testDataProvider.WriteOutAndReadBack(wb);

            Assert.AreEqual(0, wbr.GetSheetIndex("Sheet 6"));
            Assert.AreEqual(1, wbr.GetSheetIndex("Sheet 0"));
            Assert.AreEqual(2, wbr.GetSheetIndex("Sheet 2"));
            Assert.AreEqual(3, wbr.GetSheetIndex("Sheet 4"));
            Assert.AreEqual(4, wbr.GetSheetIndex("Sheet 5"));
            Assert.AreEqual(5, wbr.GetSheetIndex("Sheet 7"));
            Assert.AreEqual(6, wbr.GetSheetIndex("Sheet 3"));
            Assert.AreEqual(7, wbr.GetSheetIndex("Sheet 8"));
            Assert.AreEqual(8, wbr.GetSheetIndex("Sheet 9"));
            Assert.AreEqual(9, wbr.GetSheetIndex("Sheet 1"));

            // Now get the index by the sheet, not the name
            for (int i = 0; i < 10; i++)
            {
                ISheet s = wbr.GetSheetAt(i);
                Assert.AreEqual(i, wbr.GetSheetIndex(s));
            }
        }
        [TestMethod]
        public void TestCloneSheet()
        {
            IWorkbook book = _testDataProvider.CreateWorkbook();
            ISheet sheet = book.CreateSheet("TEST");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Test");
            sheet.CreateRow(1).CreateCell(0).SetCellValue(36.6);
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2));
            sheet.AddMergedRegion(new CellRangeAddress(1, 2, 0, 2));
            Assert.IsTrue(sheet.IsSelected);

            ISheet ClonedSheet = book.CloneSheet(0);
            Assert.AreEqual("TEST (2)", ClonedSheet.SheetName);
            Assert.AreEqual(2, ClonedSheet.PhysicalNumberOfRows);
            Assert.AreEqual(2, ClonedSheet.NumMergedRegions);
            Assert.IsFalse(ClonedSheet.IsSelected);

            //Cloned sheet is a deep copy, Adding rows in the original does not affect the clone
            sheet.CreateRow(2).CreateCell(0).SetCellValue(1);
            sheet.AddMergedRegion(new CellRangeAddress(0, 2, 0, 2));
            Assert.AreEqual(2, ClonedSheet.PhysicalNumberOfRows);
            Assert.AreEqual(2, ClonedSheet.PhysicalNumberOfRows);

            ClonedSheet.CreateRow(2).CreateCell(0).SetCellValue(1);
            ClonedSheet.AddMergedRegion(new CellRangeAddress(0, 2, 0, 2));
            Assert.AreEqual(3, ClonedSheet.PhysicalNumberOfRows);
            Assert.AreEqual(3, ClonedSheet.PhysicalNumberOfRows);

        }
        [TestMethod]
        public void TestParentReferences()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            Assert.AreSame(workbook, sheet.Workbook);

            IRow row = sheet.CreateRow(0);
            Assert.AreSame(sheet, row.Sheet);

            ICell cell = row.CreateCell(1);
            Assert.AreSame(sheet, cell.Sheet);
            Assert.AreSame(row, cell.Row);

            workbook = _testDataProvider.WriteOutAndReadBack(workbook);
            sheet = workbook.GetSheetAt(0);
            Assert.AreSame(workbook, sheet.Workbook);

            row = sheet.GetRow(0);
            Assert.AreSame(sheet, row.Sheet);

            cell = row.GetCell(1);
            Assert.AreSame(sheet, cell.Sheet);
            Assert.AreSame(row, cell.Row);
        }
        [TestMethod]
        public void TestSetRepeatingRowsAnsColumns()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            wb.SetRepeatingRowsAndColumns(wb.GetSheetIndex(sheet1), 0, 0, 0, 3);

            //must handle sheets with quotas, see Bugzilla #47294
            ISheet sheet2 = wb.CreateSheet("My' Sheet");
            wb.SetRepeatingRowsAndColumns(wb.GetSheetIndex(sheet2), 0, 0, 0, 3);
        }

        /**
         * Tests that all of the unicode capable string fields can be Set, written and then read back
         */
        [TestMethod]
        public void TestUnicodeInAll()
        {
            // This Test depends on the american culture.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            CreationHelper factory = wb.GetCreationHelper();
            //Create a unicode dataformat (Contains euro symbol)
            IDataFormat df = wb.CreateDataFormat();
            String formatStr = "_([$\u20ac-2]\\\\\\ * #,##0.00_);_([$\u20ac-2]\\\\\\ * \\\\\\(#,##0.00\\\\\\);_([$\u20ac-2]\\\\\\ *\\\"\\-\\\\\"??_);_(@_)";
            short fmt = df.GetFormat(formatStr);

            //Create a unicode sheet name (euro symbol)
            ISheet s = wb.CreateSheet("\u20ac");

            //Set a unicode header (you guessed it the euro symbol)
            IHeader h = s.Header;
            h.Center=("\u20ac");
            h.Left=("\u20ac");
            h.Right=("\u20ac");

            //Set a unicode footer
            IFooter f = s.Footer;
            f.Center=("\u20ac");
            f.Left=("\u20ac");
            f.Right=("\u20ac");

            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(1);
            c.SetCellValue(12.34);
            c.CellStyle.DataFormat=(fmt);

            ICell c2 = r.CreateCell(2); // TODO - c2 unused but changing next line ('c'->'c2') causes test to Assert.Fail
            c.SetCellValue(factory.CreateRichTextString("\u20ac"));

            ICell c3 = r.CreateCell(3);
            String formulaString = "TEXT(12.34,\"\u20ac###,##\")";
            c3.CellFormula = (formulaString);

            wb = _testDataProvider.WriteOutAndReadBack(wb);

            //Test the sheetname
            s = wb.GetSheet("\u20ac");
            Assert.IsNotNull(s);

            //Test the header
            h = s.Header;
            Assert.AreEqual(h.Center, "\u20ac");
            Assert.AreEqual(h.Left, "\u20ac");
            Assert.AreEqual(h.Right, "\u20ac");

            //Test the footer
            f = s.Footer;
            Assert.AreEqual(f.Center, "\u20ac");
            Assert.AreEqual(f.Left, "\u20ac");
            Assert.AreEqual(f.Right, "\u20ac");

            //Test the dataformat
            r = s.GetRow(0);
            c = r.GetCell(1);
            df = wb.CreateDataFormat();
            Assert.AreEqual(formatStr, df.GetFormat(c.CellStyle.DataFormat));

            //Test the cell string value
            c2 = r.GetCell(2);
            Assert.AreEqual(c.RichStringCellValue.String, "\u20ac");

            //Test the cell formula
            c3 = r.GetCell(3);
            Assert.AreEqual(c3.CellFormula, formulaString);
        }
    }
}
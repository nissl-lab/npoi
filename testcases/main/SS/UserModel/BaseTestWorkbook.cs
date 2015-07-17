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
    using NUnit.Framework;
    using NPOI.SS.Util;
    using TestCases.SS;
    using NPOI.SS.UserModel;
    using System.Text;

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public abstract class BaseTestWorkbook
    {

        private ITestDataProvider _testDataProvider;
        public BaseTestWorkbook()
        {
            _testDataProvider = TestCases.HSSF.HSSFITestDataProvider.Instance;
        }
        protected BaseTestWorkbook(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }

        [Test]
        public void TestCreateSheet()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, wb.NumberOfSheets);

            //getting a sheet by invalid index or non-existing name
            Assert.IsNull(wb.GetSheet("Sheet1"));
            try
            {
                wb.GetSheetAt(0);
                Assert.Fail("should have thrown exceptiuon due to invalid sheet index");
            }
            catch (ArgumentException ex)
            {
                // expected during successful Test
                // no negative index in the range message
                Assert.IsFalse(ex.Message.Contains("-1"));
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
                throw new AssertionException("Identified bug 44892");
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
                // expected during successful Test
                Assert.AreEqual("The workbook already contains a sheet of this name", e.Message);
            }

            //names cannot be blank or contain any of /\*?[]
            String[] invalidNames = {"", "Sheet/", "Sheet\\",
                "Sheet?", "Sheet*", "Sheet[", "Sheet]", "'Sheet'",
                "My:Sheet"};
            foreach (String sheetName in invalidNames)
            {
                try
                {
                    wb.CreateSheet(sheetName);
                    Assert.Fail("should have thrown exception due to invalid sheet name: " + sheetName);
                }
                catch (ArgumentException)
                {
                    // expected during successful Test
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
                // expected during successful Test
            }

            //try to assign an invalid name to the 2nd sheet
            try
            {
                wb.CreateSheet(null);
                Assert.Fail("should have thrown exceptiuon due to invalid sheet name");
            }
            catch (ArgumentException)
            {
                // expected during successful Test
            }

            try
            {
                wb.SetSheetName(2, null);

                Assert.Fail("should have thrown exceptiuon due to invalid sheet name");
            }
            catch (ArgumentException)
            {
                // expected during successful Test
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

        /**
         * POI allows creating sheets with names longer than 31 characters.
         *
         * Excel opens files with long sheet names without error or warning.
         * However, long sheet names are silently tRuncated to 31 chars.  In order to
         * avoid funny duplicate sheet name errors, POI enforces uniqueness on only the first 31 chars.
         * but for the purpose of uniqueness long sheet names are silently tRuncated to 31 chars.
         */
        [Test]
        public void TestCreateSheetWithLongNames()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            String sheetName1 = "My very long sheet name which is longer than 31 chars";
            String tRuncatedSheetName1 = sheetName1.Substring(0, 31);
            ISheet sh1 = wb.CreateSheet(sheetName1);
            Assert.AreEqual(tRuncatedSheetName1, sh1.SheetName);
            Assert.AreSame(sh1, wb.GetSheet(tRuncatedSheetName1));
            // now via wb.SetSheetName
            wb.SetSheetName(0, sheetName1);
            Assert.AreEqual(tRuncatedSheetName1, sh1.SheetName);
            Assert.AreSame(sh1, wb.GetSheet(tRuncatedSheetName1));

            String sheetName2 = "My very long sheet name which is longer than 31 chars " +
                    "and sheetName2.Substring(0, 31) == sheetName1.Substring(0, 31)";
            try
            {
                ISheet sh2 = wb.CreateSheet(sheetName2);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                Assert.AreEqual("The workbook already contains a sheet of this name", e.Message);
            }

            String sheetName3 = "POI allows creating sheets with names longer than 31 characters";
            String tRuncatedSheetName3 = sheetName3.Substring(0, 31);
            ISheet sh3 = wb.CreateSheet(sheetName3);
            Assert.AreEqual(tRuncatedSheetName3, sh3.SheetName);
            Assert.AreSame(sh3, wb.GetSheet(tRuncatedSheetName3));

            //serialize and read again
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            Assert.AreEqual(2, wb.NumberOfSheets);
            Assert.AreEqual(0, wb.GetSheetIndex(tRuncatedSheetName1));
            Assert.AreEqual(1, wb.GetSheetIndex(tRuncatedSheetName3));
        }

        [Test]
        public void TestRemoveSheetAt()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            try
            {
                workbook.CreateSheet("sheet1");
                workbook.CreateSheet("sheet2");
                workbook.CreateSheet("sheet3");
                Assert.AreEqual(3, workbook.NumberOfSheets);
                Assert.AreEqual(0, workbook.ActiveSheetIndex);
                workbook.RemoveSheetAt(1);
                Assert.AreEqual(2, workbook.NumberOfSheets);
                Assert.AreEqual("sheet3", workbook.GetSheetName(1));
                Assert.AreEqual(0, workbook.ActiveSheetIndex);
                workbook.RemoveSheetAt(0);
                Assert.AreEqual(1, workbook.NumberOfSheets);
                Assert.AreEqual("sheet3", workbook.GetSheetName(0));
                Assert.AreEqual(0, workbook.ActiveSheetIndex);
                workbook.RemoveSheetAt(0);
                Assert.AreEqual(0, workbook.NumberOfSheets);
                Assert.AreEqual(0, workbook.ActiveSheetIndex);

                //re-create the sheets
                workbook.CreateSheet("sheet1");
                workbook.CreateSheet("sheet2");
                workbook.CreateSheet("sheet3");
                Assert.AreEqual(3, workbook.NumberOfSheets);

                workbook.CreateSheet("sheet4");
                Assert.AreEqual(4, workbook.NumberOfSheets);

                Assert.AreEqual(0, workbook.ActiveSheetIndex);
                workbook.SetActiveSheet(2);
                Assert.AreEqual(2, workbook.ActiveSheetIndex);

                workbook.RemoveSheetAt(2);
                Assert.AreEqual(2, workbook.ActiveSheetIndex);

                workbook.RemoveSheetAt(1);
                Assert.AreEqual(1, workbook.ActiveSheetIndex);

                workbook.RemoveSheetAt(0);
                Assert.AreEqual(0, workbook.ActiveSheetIndex);

                workbook.RemoveSheetAt(0);
                Assert.AreEqual(0, workbook.ActiveSheetIndex);
            }
            finally
            {
                workbook.Close();
            }

        }

        [Test]
        public void TestDefaultValues()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, b.ActiveSheetIndex);
            Assert.AreEqual(0, b.FirstVisibleTab);
            Assert.AreEqual(0, b.NumberOfNames);
            Assert.AreEqual(0, b.NumberOfSheets);
        }

        [Test]
        public void TestSheetSelection()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            b.CreateSheet("Sheet One");
            b.CreateSheet("Sheet Two");
            b.SetActiveSheet(1);
            b.SetSelectedTab(1);
            b.FirstVisibleTab = (/*setter*/1);
            Assert.AreEqual(1, b.ActiveSheetIndex);
            Assert.AreEqual(1, b.FirstVisibleTab);
        }

        [Test]
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

        [Test]
        public void TestGetSetActiveSheet()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, workbook.ActiveSheetIndex);

            workbook.CreateSheet("sheet1");
            workbook.CreateSheet("sheet2");
            workbook.CreateSheet("sheet3");
            // Set second sheet
            workbook.SetActiveSheet(1);
            // Test if second sheet is Set up
            Assert.AreEqual(1, workbook.ActiveSheetIndex);

            workbook.SetActiveSheet(0);
            // Test if second sheet is Set up
            Assert.AreEqual(0, workbook.ActiveSheetIndex);
        }

        [Test]
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

            // check active sheet
            Assert.AreEqual(0, wb.ActiveSheetIndex);

            // Change
            wb.SetSheetOrder("Sheet 6", 0);
            Assert.AreEqual(1, wb.ActiveSheetIndex);
            wb.SetSheetOrder("Sheet 3", 7);
            wb.SetSheetOrder("Sheet 1", 9);
            // now the first sheet is at index 1
            Assert.AreEqual(1, wb.ActiveSheetIndex);

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

            Assert.AreEqual(1, wb.ActiveSheetIndex);

            // Now Get the index by the sheet, not the name
            for (int i = 0; i < 10; i++)
            {
                ISheet s = wbr.GetSheetAt(i);
                Assert.AreEqual(i, wbr.GetSheetIndex(s));
            }
        }

        [Test]
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

        [Test]
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

        /**
     * Test is kept to ensure stub for deprecated business method passes test.
     * 
     * @Deprecated remove this test when 
     * {@link Workbook#setRepeatingRowsAndColumns(int, int, int, int, int)} 
     * 
     */
        [Obsolete("remove this test when Workbook#setRepeatingRowsAndColumns(int, int, int, int, int) is removed ")]
        [Test]
        public void TestSetRepeatingRowsAnsColumns()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            wb.SetRepeatingRowsAndColumns(wb.GetSheetIndex(sheet1), 0, 0, 0, 3);
            Assert.AreEqual("1:4", sheet1.RepeatingRows.FormatAsString());
            Assert.AreEqual("A:A", sheet1.RepeatingColumns.FormatAsString());

            //must handle sheets with quotas, see Bugzilla #47294
            ISheet sheet2 = wb.CreateSheet("My' Sheet");
            wb.SetRepeatingRowsAndColumns(wb.GetSheetIndex(sheet2), 0, 0, 0, 3);
            Assert.AreEqual("1:4", sheet2.RepeatingRows.FormatAsString());
            Assert.AreEqual("A:A", sheet1.RepeatingColumns.FormatAsString());
        }

        /**
         * Tests that all of the unicode capable string fields can be Set, written and then read back
         */
        [Test]
        public void TestUnicodeInAll()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICreationHelper factory = wb.GetCreationHelper(/*getter*/);
            //Create a unicode dataformat (Contains euro symbol)
            IDataFormat df = wb.CreateDataFormat();
            String formatStr = "_([$\u20ac-2]\\\\\\ * #,##0.00_);_([$\u20ac-2]\\\\\\ * \\\\\\(#,##0.00\\\\\\);_([$\u20ac-2]\\\\\\ *\\\"\\-\\\\\"??_);_(@_)";
            short fmt = df.GetFormat(formatStr);

            //Create a unicode sheet name (euro symbol)
            ISheet s = wb.CreateSheet("\u20ac");

            //Set a unicode header (you guessed it the euro symbol)
            IHeader h = s.Header;
            h.Center = (/*setter*/"\u20ac");
            h.Left = (/*setter*/"\u20ac");
            h.Right = (/*setter*/"\u20ac");

            //Set a unicode footer
            IFooter f = s.Footer;
            f.Center = (/*setter*/"\u20ac");
            f.Left = (/*setter*/"\u20ac");
            f.Right = (/*setter*/"\u20ac");

            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(1);
            c.SetCellValue(12.34);
            c.CellStyle.DataFormat = (/*setter*/fmt);

            ICell c2 = r.CreateCell(2); // TODO - c2 unused but changing next line ('c'->'c2') causes Test to fail
            c.SetCellValue(factory.CreateRichTextString("\u20ac"));

            ICell c3 = r.CreateCell(3);
            String formulaString = "TEXT(12.34,\"\u20ac###,##\")";
            c3.CellFormula = (/*setter*/formulaString);

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

        private IWorkbook newSetSheetNameTestingWorkbook()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh1 = wb.CreateSheet("Worksheet");
            ISheet sh2 = wb.CreateSheet("Testing 47100");
            ISheet sh3 = wb.CreateSheet("To be Renamed");

            IName name1 = wb.CreateName();
            name1.NameName = (/*setter*/"sale_1");
            name1.RefersToFormula = (/*setter*/"Worksheet!$A$1");

            IName name2 = wb.CreateName();
            name2.NameName = (/*setter*/"sale_2");
            name2.RefersToFormula = (/*setter*/"'Testing 47100'!$A$1");

            IName name3 = wb.CreateName();
            name3.NameName = (/*setter*/"sale_3");
            name3.RefersToFormula = (/*setter*/"'Testing 47100'!$B$1");

            IName name4 = wb.CreateName();
            name4.NameName = (/*setter*/"sale_4");
            name4.RefersToFormula = (/*setter*/"'To be Renamed'!$A$3");

            sh1.CreateRow(0).CreateCell(0).CellFormula = (/*setter*/"SUM('Testing 47100'!A1:C1)");
            sh1.CreateRow(1).CreateCell(0).CellFormula = (/*setter*/"SUM('Testing 47100'!A1:C1,'To be Renamed'!A1:A5)");
            sh1.CreateRow(2).CreateCell(0).CellFormula = (/*setter*/"sale_2+sale_3+'Testing 47100'!C1");

            sh2.CreateRow(0).CreateCell(0).SetCellValue(1);
            sh2.GetRow(0).CreateCell(1).SetCellValue(2);
            sh2.GetRow(0).CreateCell(2).SetCellValue(3);

            sh3.CreateRow(0).CreateCell(0).SetCellValue(1);
            sh3.CreateRow(1).CreateCell(0).SetCellValue(2);
            sh3.CreateRow(2).CreateCell(0).SetCellValue(3);
            sh3.CreateRow(3).CreateCell(0).SetCellValue(4);
            sh3.CreateRow(4).CreateCell(0).SetCellValue(5);
            sh3.CreateRow(5).CreateCell(0).CellFormula = (/*setter*/"sale_3");
            sh3.CreateRow(6).CreateCell(0).CellFormula = (/*setter*/"'Testing 47100'!C1");

            return wb;
        }

        /**
         * Ensure that Workbook#setSheetName updates all dependent formulas and named ranges
         *
         * @see <a href="https://issues.apache.org/bugzilla/Show_bug.cgi?id=47100">Bugzilla 47100</a>
         */
        [Test]
        public void TestSetSheetName()
        {

            IWorkbook wb = newSetSheetNameTestingWorkbook();

            ISheet sh1 = wb.GetSheetAt(0);

            IName sale_2 = wb.GetNameAt(1);
            IName sale_3 = wb.GetNameAt(2);
            IName sale_4 = wb.GetNameAt(3);

            Assert.AreEqual("sale_2", sale_2.NameName);
            Assert.AreEqual("'Testing 47100'!$A$1", sale_2.RefersToFormula);
            Assert.AreEqual("sale_3", sale_3.NameName);
            Assert.AreEqual("'Testing 47100'!$B$1", sale_3.RefersToFormula);
            Assert.AreEqual("sale_4", sale_4.NameName);
            Assert.AreEqual("'To be Renamed'!$A$3", sale_4.RefersToFormula);

            IFormulaEvaluator Evaluator = wb.GetCreationHelper(/*getter*/).CreateFormulaEvaluator();

            ICell cell0 = sh1.GetRow(0).GetCell(0);
            ICell cell1 = sh1.GetRow(1).GetCell(0);
            ICell cell2 = sh1.GetRow(2).GetCell(0);

            Assert.AreEqual("SUM('Testing 47100'!A1:C1)", cell0.CellFormula);
            Assert.AreEqual("SUM('Testing 47100'!A1:C1,'To be Renamed'!A1:A5)", cell1.CellFormula);
            Assert.AreEqual("sale_2+sale_3+'Testing 47100'!C1", cell2.CellFormula);

            Assert.AreEqual(6.0, Evaluator.Evaluate(cell0).NumberValue);
            Assert.AreEqual(21.0, Evaluator.Evaluate(cell1).NumberValue);
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell2).NumberValue);

            wb.SetSheetName(1, "47100 - First");
            wb.SetSheetName(2, "47100 - Second");

            Assert.AreEqual("sale_2", sale_2.NameName);
            Assert.AreEqual("'47100 - First'!$A$1", sale_2.RefersToFormula);
            Assert.AreEqual("sale_3", sale_3.NameName);
            Assert.AreEqual("'47100 - First'!$B$1", sale_3.RefersToFormula);
            Assert.AreEqual("sale_4", sale_4.NameName);
            Assert.AreEqual("'47100 - Second'!$A$3", sale_4.RefersToFormula);

            Assert.AreEqual("SUM('47100 - First'!A1:C1)", cell0.CellFormula);
            Assert.AreEqual("SUM('47100 - First'!A1:C1,'47100 - Second'!A1:A5)", cell1.CellFormula);
            Assert.AreEqual("sale_2+sale_3+'47100 - First'!C1", cell2.CellFormula);

            Evaluator.ClearAllCachedResultValues();
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell0).NumberValue);
            Assert.AreEqual(21.0, Evaluator.Evaluate(cell1).NumberValue);
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell2).NumberValue);

            wb = _testDataProvider.WriteOutAndReadBack(wb);

            sh1 = wb.GetSheetAt(0);

            sale_2 = wb.GetNameAt(1);
            sale_3 = wb.GetNameAt(2);
            sale_4 = wb.GetNameAt(3);

            cell0 = sh1.GetRow(0).GetCell(0);
            cell1 = sh1.GetRow(1).GetCell(0);
            cell2 = sh1.GetRow(2).GetCell(0);

            Assert.AreEqual("sale_2", sale_2.NameName);
            Assert.AreEqual("'47100 - First'!$A$1", sale_2.RefersToFormula);
            Assert.AreEqual("sale_3", sale_3.NameName);
            Assert.AreEqual("'47100 - First'!$B$1", sale_3.RefersToFormula);
            Assert.AreEqual("sale_4", sale_4.NameName);
            Assert.AreEqual("'47100 - Second'!$A$3", sale_4.RefersToFormula);

            Assert.AreEqual("SUM('47100 - First'!A1:C1)", cell0.CellFormula);
            Assert.AreEqual("SUM('47100 - First'!A1:C1,'47100 - Second'!A1:A5)", cell1.CellFormula);
            Assert.AreEqual("sale_2+sale_3+'47100 - First'!C1", cell2.CellFormula);

            Evaluator = wb.GetCreationHelper(/*getter*/).CreateFormulaEvaluator();
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell0).NumberValue);
            Assert.AreEqual(21.0, Evaluator.Evaluate(cell1).NumberValue);
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell2).NumberValue);
        }

        public void ChangeSheetNameWithSharedFormulas(String sampleFile)
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(sampleFile);

            IFormulaEvaluator Evaluator = wb.GetCreationHelper(/*getter*/).CreateFormulaEvaluator();

            ISheet sheet = wb.GetSheetAt(0);

            for (int rownum = 1; rownum <= 40; rownum++)
            {
                ICell cellA = sheet.GetRow(1).GetCell(0);
                ICell cellB = sheet.GetRow(1).GetCell(1);

                Assert.AreEqual(cellB.StringCellValue, Evaluator.Evaluate(cellA).StringValue);
            }

            wb.SetSheetName(0, "Renamed by POI");
            Evaluator.ClearAllCachedResultValues();

            for (int rownum = 1; rownum <= 40; rownum++)
            {
                ICell cellA = sheet.GetRow(1).GetCell(0);
                ICell cellB = sheet.GetRow(1).GetCell(1);

                Assert.AreEqual(cellB.StringCellValue, Evaluator.Evaluate(cellA).StringValue);
            }
        }

        protected void assertSheetOrder(IWorkbook wb, params String[] sheets)
        {
            StringBuilder sheetNames = new StringBuilder();
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                sheetNames.Append(wb.GetSheetAt(i).SheetName).Append(",");
            }
            Assert.AreEqual(sheets.Length, wb.NumberOfSheets, "Had: " + sheetNames.ToString());
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                Assert.AreEqual(sheets[i], wb.GetSheetAt(i).SheetName, "Had: " + sheetNames.ToString());
            }
        }
    }

}
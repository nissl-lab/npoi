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
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NUnit.Framework;
    using System;
    using System.Collections;
    using System.IO;
    using System.Text;
    using TestCases.HSSF;
    using TestCases.SS;
    using TestCases.Util;

    /**
     * @author Yegor Kozlov
     */
    public abstract class BaseTestWorkbook : POITestCase
    {

        protected ITestDataProvider _testDataProvider;

        protected BaseTestWorkbook(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }

        [Test]
        public void SheetIterator_forEach()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet0");
            wb.CreateSheet("Sheet1");
            wb.CreateSheet("Sheet2");
            int i = 0;
            foreach (ISheet sh in wb)
            {
                Assert.AreEqual("Sheet" + i, sh.SheetName);
                i++;
            }
            wb.Close();
        }

        [Test]
        public void SheetIterator_sheetsReordered()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet0");
            wb.CreateSheet("Sheet1");
            wb.CreateSheet("Sheet2");

            IEnumerator it = wb.GetEnumerator();
            it.MoveNext();
            wb.SetSheetOrder("Sheet2", 1);

            // Iterator order should be fixed when iterator is created
            try
            {
                it = wb.GetEnumerator();
                it.MoveNext();
                it.MoveNext();
                it.MoveNext();
                Assert.AreEqual("Sheet1", (it.Current as ISheet).SheetName);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void SheetIterator_sheetRemoved()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet0");
            wb.CreateSheet("Sheet1");
            wb.CreateSheet("Sheet2");

            IEnumerator it = wb.GetEnumerator();
            wb.RemoveSheetAt(1);

            // Iterator order should be fixed when iterator is created
            try
            {
                it = wb.GetEnumerator();
                it.MoveNext();
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void SheetIterator_remove()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet0");

            //Iterator<Sheet> it = wb.sheetIterator();
            IEnumerator it = wb.GetEnumerator();
            it.MoveNext(); //Sheet0
            try
            {
                //it.remove();
            }
            finally
            {
                wb.Close();
            }
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
                Assert.Fail("Identified bug 44892");
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
                Assert.AreEqual("The workbook already contains a sheet named 'sHeeT3'", e.Message);
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
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb);
            wb.Close();
            Assert.AreEqual(3, wb2.NumberOfSheets);
            Assert.AreEqual(0, wb2.GetSheetIndex("sheet0"));
            Assert.AreEqual(1, wb2.GetSheetIndex("sheet1"));
            Assert.AreEqual(2, wb2.GetSheetIndex("I Changed!"));
            wb.Close();
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
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();

            String sheetName1 = "My very long sheet name which is longer than 31 chars";
            try
            {
                ISheet sh1 = wb1.CreateSheet(sheetName1);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException ex)
            {
                Assert.IsTrue(ex.Message.StartsWith("sheetName 'My very long sheet name which is longer than 31 chars' is invalid"));
            }
            try
            {
                 wb1.CreateSheet("test");
                // now via wb.SetSheetName
                wb1.SetSheetName(0, sheetName1);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException ex)
            {
                Assert.IsTrue(ex.Message.StartsWith("sheetName 'My very long sheet name which is longer than 31 chars' is invalid"));
            }
            wb1.Close();
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
            b.Close();
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
            b.Close();
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
            workbook.Close();
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
            workbook.Close();
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
            wb.Close();

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
            wbr.Close();
        }

        [Test]
        public virtual void CloneSheet()
        {
            IWorkbook book = _testDataProvider.CreateWorkbook();
            ISheet sheet = book.CreateSheet("TEST");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Test");
            sheet.CreateRow(1).CreateCell(0).SetCellValue(36.6);
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2));
            sheet.AddMergedRegion(new CellRangeAddress(2, 3, 0, 2));
            Assert.IsTrue(sheet.IsSelected);

            ISheet ClonedSheet = book.CloneSheet(0);
            Assert.AreEqual("TEST (2)", ClonedSheet.SheetName);
            Assert.AreEqual(2, ClonedSheet.PhysicalNumberOfRows);
            Assert.AreEqual(2, ClonedSheet.NumMergedRegions);
            Assert.IsFalse(ClonedSheet.IsSelected);

            //Cloned sheet is a deep copy, Adding rows or merged regions in the original does not affect the clone
            sheet.CreateRow(2).CreateCell(0).SetCellValue(1);
            sheet.AddMergedRegion(new CellRangeAddress(4, 5, 0, 2));
            Assert.AreEqual(2, ClonedSheet.PhysicalNumberOfRows);
            Assert.AreEqual(2, ClonedSheet.NumMergedRegions);

            ClonedSheet.CreateRow(2).CreateCell(0).SetCellValue(1);
            ClonedSheet.AddMergedRegion(new CellRangeAddress(6, 7, 0, 2));
            Assert.AreEqual(3, ClonedSheet.PhysicalNumberOfRows);
            Assert.AreEqual(3, ClonedSheet.NumMergedRegions);
            book.Close();
        }

        [Test]
        public void TestParentReferences()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();
            Assert.AreSame(wb1, sheet.Workbook);

            IRow row = sheet.CreateRow(0);
            Assert.AreSame(sheet, row.Sheet);

            ICell cell = row.CreateCell(1);
            Assert.AreSame(sheet, cell.Sheet);
            Assert.AreSame(row, cell.Row);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            Assert.AreSame(wb2, sheet.Workbook);

            row = sheet.GetRow(0);
            Assert.AreSame(sheet, row.Sheet);

            cell = row.GetCell(1);
            Assert.AreSame(sheet, cell.Sheet);
            Assert.AreSame(row, cell.Row);
            wb2.Close();
        }

        /**
         * Test to validate that replacement for removed setRepeatingRowsAnsColumns() methods
         * is still working correctly 
         */
        [Test]
        public void SetRepeatingRowsAnsColumns()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            CellRangeAddress cra = new CellRangeAddress(0, 3, 0, 0);
            String expRows = "1:4", expCols = "A:A";


            ISheet sheet1 = wb.CreateSheet();
            sheet1.RepeatingRows = (cra);
            sheet1.RepeatingColumns = (cra);
            Assert.AreEqual(expRows, sheet1.RepeatingRows.FormatAsString());
            Assert.AreEqual(expCols, sheet1.RepeatingColumns.FormatAsString());

            //must handle sheets with quotas, see Bugzilla #47294
            ISheet sheet2 = wb.CreateSheet("My' Sheet");
            sheet2.RepeatingRows = (cra);
            sheet2.RepeatingColumns = (cra);
            Assert.AreEqual(expRows, sheet2.RepeatingRows.FormatAsString());
            Assert.AreEqual(expCols, sheet2.RepeatingColumns.FormatAsString());
            wb.Close();
        }

        /**
         * Tests that all of the unicode capable string fields can be Set, written and then read back
         */
        [Test]
        public void TestUnicodeInAll()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ICreationHelper factory = wb1.GetCreationHelper(/*getter*/);
            //Create a unicode dataformat (Contains euro symbol)
            IDataFormat df = wb1.CreateDataFormat();
            String formatStr = "_([$\u20ac-2]\\\\\\ * #,##0.00_);_([$\u20ac-2]\\\\\\ * \\\\\\(#,##0.00\\\\\\);_([$\u20ac-2]\\\\\\ *\\\"\\-\\\\\"??_);_(@_)";
            short fmt = df.GetFormat(formatStr);

            //Create a unicode sheet name (euro symbol)
            ISheet s = wb1.CreateSheet("\u20ac");

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

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            //Test the sheetname
            s = wb2.GetSheet("\u20ac");
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
            df = wb2.CreateDataFormat();
            Assert.AreEqual(formatStr, df.GetFormat(c.CellStyle.DataFormat));

            //Test the cell string value
            c2 = r.GetCell(2);
            Assert.AreEqual(c.RichStringCellValue.String, "\u20ac");

            //Test the cell formula
            c3 = r.GetCell(3);
            Assert.AreEqual(c3.CellFormula, formulaString);

            wb2.Close();
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
        public virtual void TestSetSheetName()
        {

            IWorkbook wb1 = newSetSheetNameTestingWorkbook();

            ISheet sh1 = wb1.GetSheetAt(0);

            IName sale_2 = wb1.GetNameAt(1);
            IName sale_3 = wb1.GetNameAt(2);
            IName sale_4 = wb1.GetNameAt(3);

            Assert.AreEqual("sale_2", sale_2.NameName);
            Assert.AreEqual("'Testing 47100'!$A$1", sale_2.RefersToFormula);
            Assert.AreEqual("sale_3", sale_3.NameName);
            Assert.AreEqual("'Testing 47100'!$B$1", sale_3.RefersToFormula);
            Assert.AreEqual("sale_4", sale_4.NameName);
            Assert.AreEqual("'To be Renamed'!$A$3", sale_4.RefersToFormula);

            IFormulaEvaluator Evaluator = wb1.GetCreationHelper(/*getter*/).CreateFormulaEvaluator();

            ICell cell0 = sh1.GetRow(0).GetCell(0);
            ICell cell1 = sh1.GetRow(1).GetCell(0);
            ICell cell2 = sh1.GetRow(2).GetCell(0);

            Assert.AreEqual("SUM('Testing 47100'!A1:C1)", cell0.CellFormula);
            Assert.AreEqual("SUM('Testing 47100'!A1:C1,'To be Renamed'!A1:A5)", cell1.CellFormula);
            Assert.AreEqual("sale_2+sale_3+'Testing 47100'!C1", cell2.CellFormula);

            Assert.AreEqual(6.0, Evaluator.Evaluate(cell0).NumberValue);
            Assert.AreEqual(21.0, Evaluator.Evaluate(cell1).NumberValue);
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell2).NumberValue);

            wb1.SetSheetName(1, "47100 - First");
            wb1.SetSheetName(2, "47100 - Second");

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

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sh1 = wb2.GetSheetAt(0);

            sale_2 = wb2.GetNameAt(1);
            sale_3 = wb2.GetNameAt(2);
            sale_4 = wb2.GetNameAt(3);

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

            Evaluator = wb2.GetCreationHelper(/*getter*/).CreateFormulaEvaluator();
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell0).NumberValue);
            Assert.AreEqual(21.0, Evaluator.Evaluate(cell1).NumberValue);
            Assert.AreEqual(6.0, Evaluator.Evaluate(cell2).NumberValue);
            wb2.Close();
        }

        protected void ChangeSheetNameWithSharedFormulas(String sampleFile)
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
            wb.Close();
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


        

        [Test]
        public void Test58499()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            for (int i = 0; i < 900; i++)
            {
                IRow r = sheet.CreateRow(i);
                ICell c = r.CreateCell(0);
                ICellStyle cs = workbook.CreateCellStyle();
                c.CellStyle = (cs);
                c.SetCellValue("AAA");
            }
            OutputStream os = new NullOutputStream();
            try
            {
                workbook.Write(os, false);
            }
            finally
            {
                os.Close();
            }
            //workbook.dispose();
            workbook.Close();
        }


        [Test]
        public void WindowOneDefaults()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            try
            {
                Assert.AreEqual(b.ActiveSheetIndex, 0);
                Assert.AreEqual(b.FirstVisibleTab, 0);
            }
            catch (NullReferenceException)
            {
                Assert.Fail("WindowOneRecord in Workbook is probably not initialized");
            }

            b.Close();
        }

        [Test]
        public void GetSpreadsheetVersion()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(_testDataProvider.GetSpreadsheetVersion(), wb.SpreadsheetVersion);
            wb.Close();
        }

        protected void verifySpreadsheetVersion(SpreadsheetVersion expected)
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(expected, wb.SpreadsheetVersion);
            wb.Close();
        }

        protected static void assertCloseDoesNotModifyFile(String filename, IWorkbook wb) {
            byte[] before = HSSFTestDataSamples.GetTestDataFileContent(filename);
            wb.Close();
            byte[] after = HSSFTestDataSamples.GetTestDataFileContent(filename);
            CollectionAssert.AreEqual(before, after,
                filename + " sample file was modified as a result of closing the workbook");
        }

        [Test]
        public virtual void SheetClone()
        {
            // First up, try a simple file
            IWorkbook b = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, b.NumberOfSheets);
            b.CreateSheet("Sheet One");
            b.CreateSheet("Sheet Two");
            Assert.AreEqual(2, b.NumberOfSheets);
            b.CloneSheet(0);
            Assert.AreEqual(3, b.NumberOfSheets);
            // Now try a problem one with drawing records in it
            IWorkbook bBack = HSSFTestDataSamples.OpenSampleWorkbook("SheetWithDrawing.xls");
            Assert.AreEqual(1, bBack.NumberOfSheets);
            bBack.CloneSheet(0);
            Assert.AreEqual(2, bBack.NumberOfSheets);
            bBack.Close();
            b.Close();
        }
        [Test]
        public void GetSheetIndex()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            ISheet sheet2 = wb.CreateSheet("Sheet2");
            ISheet sheet3 = wb.CreateSheet("Sheet3");
            ISheet sheet4 = wb.CreateSheet("Sheet4");
            Assert.AreEqual(0, wb.GetSheetIndex(sheet1));
            Assert.AreEqual(1, wb.GetSheetIndex(sheet2));
            Assert.AreEqual(2, wb.GetSheetIndex(sheet3));
            Assert.AreEqual(3, wb.GetSheetIndex(sheet4));
            // remove sheets
            wb.RemoveSheetAt(0);
            wb.RemoveSheetAt(2);
            // ensure that sheets are moved up and removed sheets are not found any more
            Assert.AreEqual(-1, wb.GetSheetIndex(sheet1));
            Assert.AreEqual(0, wb.GetSheetIndex(sheet2));
            Assert.AreEqual(1, wb.GetSheetIndex(sheet3));
            Assert.AreEqual(-1, wb.GetSheetIndex(sheet4));
            wb.Close();
        }
        [Test]
        public void AddSheetTwice()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            Assert.IsNotNull(sheet1);
            try
            {
                wb.CreateSheet("Sheet1");
                Assert.Fail("Should Assert.Fail if we add the same sheet twice");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.Contains("already contains a sheet named 'Sheet1'"), e.Message);
            }
            wb.Close();
        }

        // bug 51233 and 55075: correctly size image if Added to a row with a custom height
        [Test]
        public virtual void CreateDrawing()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("Main Sheet");
            IRow row0 = sheet.CreateRow(0);
            IRow row1 = sheet.CreateRow(1);
            ICell cell = row1.CreateCell(0);
            row0.CreateCell(1);
            row1.CreateCell(0);
            row1.CreateCell(1);

            byte[] pictureData = _testDataProvider.GetTestDataFileContent("logoKarmokar4.png");

            int handle = wb.AddPicture(pictureData, PictureType.PNG);
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            ICreationHelper helper = wb.GetCreationHelper();
            IClientAnchor anchor = helper.CreateClientAnchor();
            anchor.AnchorType = (/*setter*/AnchorType.DontMoveAndResize);
            anchor.Col1 = (/*setter*/0);
            anchor.Row1 = (/*setter*/0);
            IPicture picture = Drawing.CreatePicture(anchor, handle);

            row0.HeightInPoints = (/*setter*/144);
            // Set a column width so that XSSF and SXSSF have the same width (default widths may be different otherwise)
            sheet.SetColumnWidth(0, 100 * 256);
            picture.Resize();

            // The actual dimensions don't matter as much as having XSSF and SXSSF produce the same size Drawings

            // Check Drawing height
            Assert.AreEqual(0, anchor.Row1);
            Assert.AreEqual(0, anchor.Row2);
            Assert.AreEqual(0, anchor.Dy1);
            Assert.AreEqual(1609725, anchor.Dy2); //HSSF: 225

            // Check Drawing width
            Assert.AreEqual(0, anchor.Col1);
            Assert.AreEqual(0, anchor.Col2);
            Assert.AreEqual(0, anchor.Dx1);
            Assert.AreEqual(1114425, anchor.Dx2); //HSSF: 171

            bool WriteOut = false;
            if (WriteOut)
            {
                string ext = "." + _testDataProvider.StandardFileNameExtension;
                string prefix = wb.GetType().Name + "-CreateDrawing";
                FileInfo f = TempFile.CreateTempFile(prefix, ext);
                FileStream out1 = new FileStream(f.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                wb.Write(out1);
                out1.Close();
            }
            wb.Close();
        }

    }

}
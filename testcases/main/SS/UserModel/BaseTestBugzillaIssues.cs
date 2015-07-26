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

    using NPOI.HSSF.Util;
    using NPOI.SS;
    using NPOI.SS.Util;
    using System.Text;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using NPOI.HSSF.UserModel;

    /**
     * A base class for bugzilla issues that can be described in terms of common ss interfaces.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class BaseTestBugzillaIssues
    {

        private ITestDataProvider _testDataProvider;
        public BaseTestBugzillaIssues()
        {
            _testDataProvider = TestCases.HSSF.HSSFITestDataProvider.Instance;
        }
        protected BaseTestBugzillaIssues(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }

        public static void assertAlmostEquals(double expected, double actual, float factor)
        {
            double diff = Math.Abs(expected - actual);
            double fuzz = expected * factor;
            if (diff > fuzz)
                Assert.Fail(actual + " not within " + fuzz + " of " + expected);
        }
        /**
         * Test writing a hyperlink
         * Open resulting sheet in Excel and check that A1 Contains a hyperlink
         *
         * Also Tests bug 15353 (problems with hyperlinks to Google)
         */
        [Test]
        public void Test23094()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            r.CreateCell(0).CellFormula = (/*setter*/"HYPERLINK(\"http://jakarta.apache.org\",\"Jakarta\")");
            r.CreateCell(1).CellFormula = (/*setter*/"HYPERLINK(\"http://google.com\",\"Google\")");

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            r = wb.GetSheetAt(0).GetRow(0);

            ICell cell_0 = r.GetCell(0);
            Assert.AreEqual("HYPERLINK(\"http://jakarta.apache.org\",\"Jakarta\")", cell_0.CellFormula);
            ICell cell_1 = r.GetCell(1);
            Assert.AreEqual("HYPERLINK(\"http://google.com\",\"Google\")", cell_1.CellFormula);
        }

        /**
         * Test writing a file with large number of unique strings,
         * open resulting file in Excel to check results!
         * @param  num the number of strings to generate
         */
        public void BaseTest15375(int num)
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            ICreationHelper factory = wb.GetCreationHelper();

            String tmp1 = null;
            String tmp2 = null;
            String tmp3 = null;

            for (int i = 0; i < num; i++)
            {
                tmp1 = "Test1" + i;
                tmp2 = "Test2" + i;
                tmp3 = "Test3" + i;

                IRow row = sheet.CreateRow(i);

                ICell cell = row.CreateCell(0);
                cell.SetCellValue(factory.CreateRichTextString(tmp1));
                cell = row.CreateCell(1);
                cell.SetCellValue(factory.CreateRichTextString(tmp2));
                cell = row.CreateCell(2);
                cell.SetCellValue(factory.CreateRichTextString(tmp3));
            }
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            for (int i = 0; i < num; i++)
            {
                tmp1 = "Test1" + i;
                tmp2 = "Test2" + i;
                tmp3 = "Test3" + i;

                IRow row = sheet.GetRow(i);

                Assert.AreEqual(tmp1, row.GetCell(0).StringCellValue);
                Assert.AreEqual(tmp2, row.GetCell(1).StringCellValue);
                Assert.AreEqual(tmp3, row.GetCell(2).StringCellValue);
            }
        }

        /**
         * Merged regions were being Removed from the parent in Cloned sheets
         */
        [Test]
        public void Test22720()
        {
            IWorkbook workBook = _testDataProvider.CreateWorkbook();
            workBook.CreateSheet("TEST");
            ISheet template = workBook.GetSheetAt(0);

            template.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2));
            template.AddMergedRegion(new CellRangeAddress(1, 2, 0, 2));

            ISheet clone = workBook.CloneSheet(0);
            int originalMerged = template.NumMergedRegions;
            Assert.AreEqual(2, originalMerged, "2 merged regions");

            //remove merged regions from clone
            for (int i = template.NumMergedRegions - 1; i >= 0; i--)
            {
                clone.RemoveMergedRegion(i);
            }

            Assert.AreEqual(originalMerged, template.NumMergedRegions, "Original Sheet's Merged Regions were Removed");
            //check if template's merged regions are OK
            if (template.NumMergedRegions > 0)
            {
                // fetch the first merged region...EXCEPTION OCCURS HERE
                template.GetMergedRegion(0);
            }
            //make sure we dont exception

        }
        [Test]
        public void Test28031()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            String formulaText =
                "IF(ROUND(A2*B2*C2,2)>ROUND(B2*D2,2),ROUND(A2*B2*C2,2),ROUND(B2*D2,2))";
            cell.CellFormula = (/*setter*/formulaText);

            Assert.AreEqual(formulaText, cell.CellFormula);
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            cell = wb.GetSheetAt(0).GetRow(0).GetCell(0);
            Assert.AreEqual("IF(ROUND(A2*B2*C2,2)>ROUND(B2*D2,2),ROUND(A2*B2*C2,2),ROUND(B2*D2,2))", cell.CellFormula);
        }

        /**
         * Bug 21334: "File error: data may have been lost" with a file
         * that Contains macros and this formula:
         * {=SUM(IF(FREQUENCY(IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),""),IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),""))>0,1))}
         */
        [Test]
        public void Test21334()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            ICell cell = sh.CreateRow(0).CreateCell(0);
            String formula = "SUM(IF(FREQUENCY(IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),\"\"),IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),\"\"))>0,1))";
            cell.CellFormula = (/*setter*/formula);

            IWorkbook wb_sv = _testDataProvider.WriteOutAndReadBack(wb);
            ICell cell_sv = wb_sv.GetSheetAt(0).GetRow(0).GetCell(0);
            Assert.AreEqual(formula, cell_sv.CellFormula);
        }

        /** another Test for the number of unique strings issue
         *test opening the resulting file in Excel*/
        [Test]
        public void Test22568()
        {
            int r = 2000; int c = 3;

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("ExcelTest");

            int col_cnt = 0, rw_cnt = 0;

            col_cnt = c;
            rw_cnt = r;

            IRow rw;
            rw = sheet.CreateRow(0);
            //Header row
            for (int j = 0; j < col_cnt; j++)
            {
                ICell cell = rw.CreateCell(j);
                cell.SetCellValue("Col " + (j + 1));
            }

            for (int i = 1; i < rw_cnt; i++)
            {
                rw = sheet.CreateRow(i);
                for (int j = 0; j < col_cnt; j++)
                {
                    ICell cell = rw.CreateCell(j);
                    cell.SetCellValue("Row:" + (i + 1) + ",Column:" + (j + 1));
                }
            }

            sheet.DefaultColumnWidth = (/*setter*/18);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            rw = sheet.GetRow(0);
            //Header row
            for (int j = 0; j < col_cnt; j++)
            {
                ICell cell = rw.GetCell(j);
                Assert.AreEqual("Col " + (j + 1), cell.StringCellValue);
            }
            for (int i = 1; i < rw_cnt; i++)
            {
                rw = sheet.GetRow(i);
                for (int j = 0; j < col_cnt; j++)
                {
                    ICell cell = rw.GetCell(j);
                    Assert.AreEqual("Row:" + (i + 1) + ",Column:" + (j + 1), cell.StringCellValue);
                }
            }
        }

        /**
         * Bug 42448: Can't parse SUMPRODUCT(A!C7:A!C67, B8:B68) / B69
         */
        [Test]
        public void Test42448()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            cell.CellFormula = (/*setter*/"SUMPRODUCT(A!C7:A!C67, B8:B68) / B69");
            Assert.IsTrue(true, "no errors parsing formula");
        }
        [Test]
        public void Test18800()
        {
            IWorkbook book = _testDataProvider.CreateWorkbook();
            book.CreateSheet("TEST");
            ISheet sheet = book.CloneSheet(0);
            book.SetSheetName(1, "CLONE");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Test");

            book = _testDataProvider.WriteOutAndReadBack(book);
            sheet = book.GetSheet("CLONE");
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            Assert.AreEqual("Test", cell.RichStringCellValue.String);
        }

        private static void AddNewSheetWithCellsA1toD4(IWorkbook book, int sheet)
        {

            ISheet sht = book.CreateSheet("s" + sheet);
            for (int r = 0; r < 4; r++)
            {

                IRow row = sht.CreateRow(r);
                for (int c = 0; c < 4; c++)
                {

                    ICell cel = row.CreateCell(c);
                    cel.SetCellValue(sheet * 100 + r * 10 + c);
                }
            }
        }
        [Test]
        public void TestBug43093()
        {
            IWorkbook xlw = _testDataProvider.CreateWorkbook();

            AddNewSheetWithCellsA1toD4(xlw, 1);
            AddNewSheetWithCellsA1toD4(xlw, 2);
            AddNewSheetWithCellsA1toD4(xlw, 3);
            AddNewSheetWithCellsA1toD4(xlw, 4);

            ISheet s2 = xlw.GetSheet("s2");
            IRow s2r3 = s2.GetRow(3);
            ICell s2E4 = s2r3.CreateCell(4);
            s2E4.CellFormula = (/*setter*/"SUM(s3!B2:C3)");

            IFormulaEvaluator eva = xlw.GetCreationHelper().CreateFormulaEvaluator();
            double d = eva.Evaluate(s2E4).NumberValue;

            Assert.AreEqual(d, (311 + 312 + 321 + 322), 0.0000001);
        }
        [Test]
        public void TestMaxFunctionArguments_bug46729()
        {
            String[] func = { "COUNT", "AVERAGE", "MAX", "MIN", "OR", "SUBTOTAL", "SKEW" };

            SpreadsheetVersion ssVersion = _testDataProvider.GetSpreadsheetVersion();
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);

            String fmla;
            foreach (String name in func)
            {

                fmla = CreateFunction(name, 5);
                cell.CellFormula = (/*setter*/fmla);

                fmla = CreateFunction(name, ssVersion.MaxFunctionArgs);
                cell.CellFormula = (/*setter*/fmla);

                try
                {
                    fmla = CreateFunction(name, ssVersion.MaxFunctionArgs + 1);
                    cell.CellFormula = (/*setter*/fmla);
                    Assert.Fail("Expected FormulaParseException");
                }
                catch (NPOI.SS.Formula.FormulaParseException e)
                {
                    Assert.IsTrue(e.Message.StartsWith("Too many arguments to function '" + name + "'"));
                }
            }
        }

        private static String CreateFunction(String name, int maxArgs)
        {
            StringBuilder fmla = new StringBuilder();
            fmla.Append(name);
            fmla.Append("(");
            for (int i = 0; i < maxArgs; i++)
            {
                if (i > 0) fmla.Append(',');
                fmla.Append("A1");
            }
            fmla.Append(")");
            return fmla.ToString();
        }
        [Test]
        public void TestAutoSize_bug506819()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            BaseTestSheetAutosizeColumn.FixFonts(wb);
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);

            String longValue = "www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, www.hostname.com";

            cell0.SetCellValue(longValue);

            sheet.AutoSizeColumn(0);
            Assert.AreEqual(255 * 256, sheet.GetColumnWidth(0)); // maximum column width is 255 characters
            sheet.SetColumnWidth(0, sheet.GetColumnWidth(0)); // Bug 506819 reports exception at this point
        }

        /**
         * CreateFreezePane column/row order check
         */
        [Test]
        public void Test49381()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            int colSplit = 1;
            int rowSplit = 2;
            int leftmostColumn = 3;
            int topRow = 4;

            ISheet s = wb.CreateSheet();

            // Populate
            for (int rn = 0; rn <= topRow; rn++)
            {
                IRow r = s.CreateRow(rn);
                for (int cn = 0; cn < leftmostColumn; cn++)
                {
                    ICell c = r.CreateCell(cn, CellType.Numeric);
                    c.SetCellValue(100 * rn + cn);
                }
            }

            // Create the Freeze Pane
            s.CreateFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
            PaneInformation paneInfo = s.PaneInformation;

            // Check it
            Assert.AreEqual(colSplit, paneInfo.VerticalSplitPosition);
            Assert.AreEqual(rowSplit, paneInfo.HorizontalSplitPosition);
            Assert.AreEqual(leftmostColumn, paneInfo.VerticalSplitLeftColumn);
            Assert.AreEqual(topRow, paneInfo.HorizontalSplitTopRow);


            // Now a row only freezepane
            s.CreateFreezePane(0, 3);
            paneInfo = s.PaneInformation;

            Assert.AreEqual(0, paneInfo.VerticalSplitPosition);
            Assert.AreEqual(3, paneInfo.HorizontalSplitPosition);
            Assert.AreEqual(0, paneInfo.VerticalSplitLeftColumn);
            Assert.AreEqual(3, paneInfo.HorizontalSplitTopRow);

            // Now a column only freezepane
            s.CreateFreezePane(4, 0);
            paneInfo = s.PaneInformation;

            Assert.AreEqual(4, paneInfo.VerticalSplitPosition);
            Assert.AreEqual(0, paneInfo.HorizontalSplitPosition);
            Assert.AreEqual(4, paneInfo.VerticalSplitLeftColumn);
            Assert.AreEqual(0, paneInfo.HorizontalSplitTopRow);
        }

        /** 
         * Test hyperlinks
         * open resulting file in excel, and check that there is a link to Google
         */
        [Test]
        public void Bug15353()
        {
            String hyperlinkF = "HYPERLINK(\"http://google.com\",\"Google\")";

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("My sheet");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.CellFormula = (/*setter*/hyperlinkF);

            Assert.AreEqual(hyperlinkF, cell.CellFormula);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            sheet = wb.GetSheet("My Sheet");
            row = sheet.GetRow(0);
            cell = row.GetCell(0);

            Assert.AreEqual(hyperlinkF, cell.CellFormula);
        }

        /**
         * HLookup and VLookup with optional arguments 
         */
        [Test]
        public void Bug51024()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r1 = s.CreateRow(0);
            IRow r2 = s.CreateRow(1);

            r1.CreateCell(0).SetCellValue("v A1");
            r2.CreateCell(0).SetCellValue("v A2");
            r1.CreateCell(1).SetCellValue("v B1");

            ICell c = r1.CreateCell(4);

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            c.SetCellFormula("VLOOKUP(\"v A1\", A1:B2, 1)");
            Assert.AreEqual("v A1", eval.Evaluate(c).StringValue);

            c.SetCellFormula("VLOOKUP(\"v A1\", A1:B2, 1, 1)");
            Assert.AreEqual("v A1", eval.Evaluate(c).StringValue);

            c.SetCellFormula("VLOOKUP(\"v A1\", A1:B2, 1, )");
            Assert.AreEqual("v A1", eval.Evaluate(c).StringValue);


            c.SetCellFormula("HLOOKUP(\"v A1\", A1:B2, 1)");
            Assert.AreEqual("v A1", eval.Evaluate(c).StringValue);

            c.SetCellFormula("HLOOKUP(\"v A1\", A1:B2, 1, 1)");
            Assert.AreEqual("v A1", eval.Evaluate(c).StringValue);

            c.SetCellFormula("HLOOKUP(\"v A1\", A1:B2, 1, )");
            Assert.AreEqual("v A1", eval.Evaluate(c).StringValue);
        }

        [Test]
        public void Stackoverflow23114397()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IDataFormat format = wb.GetCreationHelper().CreateDataFormat();

            // How close the sizing should be, given that not all
            //  systems will have quite the same fonts on them
            float fontAccuracy = 0.22f;

            // x%
            ICellStyle iPercent = wb.CreateCellStyle();
            iPercent.DataFormat = (/*setter*/format.GetFormat("0%"));
            // x.x%
            ICellStyle d1Percent = wb.CreateCellStyle();
            d1Percent.DataFormat = (/*setter*/format.GetFormat("0.0%"));
            // x.xx%
            ICellStyle d2Percent = wb.CreateCellStyle();
            d2Percent.DataFormat = (/*setter*/format.GetFormat("0.00%"));

            ISheet s = wb.CreateSheet();
            IRow r1 = s.CreateRow(0);

            for (int i = 0; i < 3; i++)
            {
                r1.CreateCell(i, CellType.Numeric).SetCellValue(0);
            }
            for (int i = 3; i < 6; i++)
            {
                r1.CreateCell(i, CellType.Numeric).SetCellValue(1);
            }
            for (int i = 6; i < 9; i++)
            {
                r1.CreateCell(i, CellType.Numeric).SetCellValue(0.12345);
            }
            for (int i = 9; i < 12; i++)
            {
                r1.CreateCell(i, CellType.Numeric).SetCellValue(1.2345);
            }
            for (int i = 0; i < 12; i += 3)
            {
                r1.GetCell(i + 0).CellStyle = (/*setter*/iPercent);
                r1.GetCell(i + 1).CellStyle = (/*setter*/d1Percent);
                r1.GetCell(i + 2).CellStyle = (/*setter*/d2Percent);
            }
            for (int i = 0; i < 12; i++)
            {
                s.AutoSizeColumn(i);
            }

            // Check the 0(.00)% ones
            assertAlmostEquals(980, s.GetColumnWidth(0), fontAccuracy);
            assertAlmostEquals(1400, s.GetColumnWidth(1), fontAccuracy);
            assertAlmostEquals(1700, s.GetColumnWidth(2), fontAccuracy);

            // Check the 100(.00)% ones
            assertAlmostEquals(1500, s.GetColumnWidth(3), fontAccuracy);
            assertAlmostEquals(1950, s.GetColumnWidth(4), fontAccuracy);
            assertAlmostEquals(2225, s.GetColumnWidth(5), fontAccuracy);

            // Check the 12(.34)% ones
            assertAlmostEquals(1225, s.GetColumnWidth(6), fontAccuracy);
            assertAlmostEquals(1650, s.GetColumnWidth(7), fontAccuracy);
            assertAlmostEquals(1950, s.GetColumnWidth(8), fontAccuracy);

            // Check the 123(.45)% ones
            assertAlmostEquals(1500, s.GetColumnWidth(9), fontAccuracy);
            assertAlmostEquals(1950, s.GetColumnWidth(10), fontAccuracy);
            assertAlmostEquals(2225, s.GetColumnWidth(11), fontAccuracy);
        }

        /**
         * =ISNUMBER(SEARCH("AM",A1)) Evaluation 
         */
        [Test]
        public void Stackoverflow26437323()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r1 = s.CreateRow(0);
            IRow r2 = s.CreateRow(1);

            // A1 is a number
            r1.CreateCell(0).SetCellValue(1.1);
            // B1 is a string, with the wanted text in it
            r1.CreateCell(1).SetCellValue("This is text with AM in it");
            // C1 is a string, with different text
            r1.CreateCell(2).SetCellValue("This some other text");
            // D1 is a blank cell
            r1.CreateCell(3, CellType.Blank);
            // E1 is null

            // A2 will hold our test formulas
            ICell cf = r2.CreateCell(0, CellType.Formula);


            // First up, check that TRUE and ISLOGICAL both behave
            cf.CellFormula = (/*setter*/"TRUE()");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISLOGICAL(TRUE())");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISLOGICAL(4)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);


            // Now, check ISNUMBER / ISTEXT / ISNONTEXT
            cf.CellFormula = (/*setter*/"ISNUMBER(A1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNUMBER(B1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNUMBER(C1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNUMBER(D1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNUMBER(E1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);


            cf.CellFormula = (/*setter*/"ISTEXT(A1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISTEXT(B1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISTEXT(C1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISTEXT(D1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISTEXT(E1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);


            cf.CellFormula = (/*setter*/"ISNONTEXT(A1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNONTEXT(B1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNONTEXT(C1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNONTEXT(D1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.CellFormula = (/*setter*/"ISNONTEXT(E1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue); // Blank and Null the same


            // Next up, SEARCH on its own
            cf.SetCellFormula("SEARCH(\"am\", A1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(ErrorConstants.ERROR_VALUE, cf.ErrorCellValue);

            cf.SetCellFormula("SEARCH(\"am\", B1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(19, (int)cf.NumericCellValue);

            cf.SetCellFormula("SEARCH(\"am\", C1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(ErrorConstants.ERROR_VALUE, cf.ErrorCellValue);

            cf.SetCellFormula("SEARCH(\"am\", D1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(ErrorConstants.ERROR_VALUE, cf.ErrorCellValue);


            // Finally, bring it all together
            cf.SetCellFormula("ISNUMBER(SEARCH(\"am\", A1))");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.SetCellFormula("ISNUMBER(SEARCH(\"am\", B1))");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(true, cf.BooleanCellValue);

            cf.SetCellFormula("ISNUMBER(SEARCH(\"am\", C1))");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.SetCellFormula("ISNUMBER(SEARCH(\"am\", D1))");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);

            cf.SetCellFormula("ISNUMBER(SEARCH(\"am\", E1))");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(false, cf.BooleanCellValue);
        }
        private ICell EvaluateCell(IWorkbook wb, ICell c)
        {
            ISheet s = c.Sheet;
            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(c);
            return s.GetRow(c.RowIndex).GetCell(c.ColumnIndex);
        }

        /**
     * Should be able to write then read formulas with references
     *  to cells in other files, eg '[refs/airport.xls]Sheet1'!$A$2
     *  or 'http://192.168.1.2/[blank.xls]Sheet1'!$A$1 .
     * Additionally, if a reference to that file is provided, it should
     *  be possible to Evaluate them too
     * TODO Fix this to Evaluate for XSSF
     * TODO Fix this to work at all for HSSF
     */
        //    [Test]
        public void bug46670()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r1 = s.CreateRow(0);


            // References to try
            String ext = "xls";
            if (!(wb is HSSFWorkbook)) ext += "x";
            String refLocal = "'[test." + ext + "]Sheet1'!$A$2";
            String refHttp = "'[http://example.com/test." + ext + "]Sheet1'!$A$2";
            String otherCellText = "In Another Workbook";


            // Create the references
            ICell c1 = r1.CreateCell(0, CellType.Formula);
            c1.CellFormula = (/*setter*/refLocal);

            ICell c2 = r1.CreateCell(1, CellType.Formula);
            c2.CellFormula = (/*setter*/refHttp);


            // Check they were Set correctly
            Assert.AreEqual(refLocal, c1.CellFormula);
            Assert.AreEqual(refHttp, c2.CellFormula);


            // Reload, and ensure they were serialised and read correctly
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r1 = s.GetRow(0);

            c1 = r1.GetCell(0);
            c2 = r1.GetCell(1);
            Assert.AreEqual(refLocal, c1.CellFormula);
            Assert.AreEqual(refHttp, c2.CellFormula);


            // Try to Evalutate, without giving a way to Get at the other file
            try
            {
                EvaluateCell(wb, c1);
                Assert.Fail("Shouldn't be able to Evaluate without the other file");
            }
            catch (Exception) { }
            try
            {
                EvaluateCell(wb, c2);
                Assert.Fail("Shouldn't be able to Evaluate without the other file");
            }
            catch (Exception) { }


            // Set up references to the other file
            IWorkbook wb2 = _testDataProvider.CreateWorkbook();
            wb2.CreateSheet().CreateRow(1).CreateCell(0).SetCellValue(otherCellText);

            Dictionary<String, IFormulaEvaluator> evaluators = new Dictionary<String, IFormulaEvaluator>();
            evaluators.Add(refLocal, wb2.GetCreationHelper().CreateFormulaEvaluator());
            evaluators.Add(refHttp, wb2.GetCreationHelper().CreateFormulaEvaluator());

            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            Evaluator.SetupReferencedWorkbooks(/*setter*/evaluators);


            // Try to Evaluate, with the other file
            Evaluator.EvaluateFormulaCell(c1);
            Evaluator.EvaluateFormulaCell(c2);

            Assert.AreEqual(otherCellText, c1.StringCellValue);
            Assert.AreEqual(otherCellText, c2.StringCellValue);
        }

    }

}
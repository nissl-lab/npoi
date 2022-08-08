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
    using NPOI.HSSF.UserModel;
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NUnit.Framework;
    using SixLabors.Fonts;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    /**
     * A base class for bugzilla issues that can be described in terms of common ss interfaces.
     *
     * @author Yegor Kozlov
     */
    public abstract class BaseTestBugzillaIssues
    {

        private ITestDataProvider _testDataProvider;

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
            r.CreateCell(0).CellFormula = "HYPERLINK(\"http://jakarta.apache.org\",\"Jakarta\")";
            r.CreateCell(1).CellFormula = "HYPERLINK(\"http://google.com\",\"Google\")";

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

            for (int i = 0; i < num; i++)
            {
                string tmp1 = "Test1" + i;
                string tmp2 = "Test2" + i;
                string tmp3 = "Test3" + i;

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
                string tmp1 = "Test1" + i;
                string tmp2 = "Test2" + i;
                string tmp3 = "Test3" + i;

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
        public virtual void Bug22720()
        {
            IWorkbook workBook = _testDataProvider.CreateWorkbook();
            workBook.CreateSheet("TEST");
            ISheet template = workBook.GetSheetAt(0);

            template.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2));
            template.AddMergedRegion(new CellRangeAddress(2, 3, 0, 2));

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
        public void TestBug28031()
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
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("ExcelTest");

            int col_cnt = 3;
            int rw_cnt = 2000;

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
        public virtual void Bug18800()
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
        public virtual void Bug46729_testMaxFunctionArguments()
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
        public void Bug50681_TestAutoSize()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            BaseTestSheetAutosizeColumn.FixFonts(wb);
            ISheet sheet = wb.CreateSheet("Sheet1");
            _testDataProvider.TrackAllColumnsForAutosizing(sheet);
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);

            String longValue = "www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, " +
                    "www.hostname.com, www.hostname.com, www.hostname.com, www.hostname.com";

            cell0.SetCellValue(longValue);


            // autoSize will fail if required fonts are not installed, skip this test then
            IFont font = wb.GetFontAt(cell0.CellStyle.FontIndex);
            Assume.That(SheetUtil.CanComputeColumnWidth(font),
                "Cannot verify auoSizeColumn() because the necessary Fonts are not installed on this machine: " + font);

            Assert.AreEqual(0, cell0.CellStyle.Indention, "Expecting no indentation in this test");
            Assert.AreEqual(0, cell0.CellStyle.Rotation, "Expecting no rotation in this test");

            // check computing size up to a large size
            StringBuilder b = new StringBuilder();
            for (int i = 0; i < longValue.Length * 5; i++)
            {
                b.Append("w");
                Assert.IsTrue(ComputeCellWidthFixed(font, b.ToString()) > 0, "Had zero length starting at length " + i);
            }

            double widthManual = ComputeCellWidthManually(cell0, font);
            double widthBeforeCell = SheetUtil.GetCellWidth(cell0, 8, null, false);
            double widthBeforeCol = SheetUtil.GetColumnWidth(sheet, 0, false);
            String info = widthManual + "/" + widthBeforeCell + "/" + widthBeforeCol + "/" +
                SheetUtil.CanComputeColumnWidth(font) + "/" + ComputeCellWidthFixed(font, "1") + "/" + ComputeCellWidthFixed(font, "w") + "/" +
                ComputeCellWidthFixed(font, "1w") + "/" + ComputeCellWidthFixed(font, "0000") + "/" + ComputeCellWidthFixed(font, longValue);

            Assert.IsTrue(widthManual > 0, "Expected to have cell width > 0 when computing manually, but had " + info);
            Assert.IsTrue(widthBeforeCell > 0, "Expected to have cell width > 0 BEFORE auto-size, but had " + info);
            Assert.IsTrue(widthBeforeCol > 0, "Expected to have column width > 0 BEFORE auto-size, but had " + info);

            sheet.AutoSizeColumn(0);

            double width = SheetUtil.GetColumnWidth(sheet, 0, false);
            Assert.IsTrue(width > 0, "Expected to have column width > 0 AFTER auto-size, but had " + width);
            width = SheetUtil.GetCellWidth(cell0, 8, null, false);
            Assert.IsTrue(width > 0, "Expected to have cell width > 0 AFTER auto-size, but had " + width);


            Assert.AreEqual(255 * 256, sheet.GetColumnWidth(0)); // maximum column width is 255 characters
            sheet.SetColumnWidth(0, sheet.GetColumnWidth(0)); // Bug 50681 reports exception at this point
        }

        [Test]
        public void Bug51622_testAutoSizeShouldRecognizeLeadingSpaces()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            BaseTestSheetAutosizeColumn.FixFonts(wb);
            ISheet sheet = wb.CreateSheet();
            _testDataProvider.TrackAllColumnsForAutosizing(sheet);
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            ICell cell1 = row.CreateCell(1);
            ICell cell2 = row.CreateCell(2);

            cell0.SetCellValue("Test Column AutoSize");
            cell1.SetCellValue("         Test Column AutoSize");
            cell2.SetCellValue("Test Column AutoSize         ");

            sheet.AutoSizeColumn(0);
            sheet.AutoSizeColumn(1);
            sheet.AutoSizeColumn(2);

            int noWhitespaceColWidth = sheet.GetColumnWidth(0);
            int leadingWhitespaceColWidth = sheet.GetColumnWidth(1);
            int trailingWhitespaceColWidth = sheet.GetColumnWidth(2);

            // Based on the amount of text and whitespace used, and the default font
            // assume that the cell with whitespace should be at least 20% wider than
            // the cell without whitespace. This number is arbitrary, but should be large
            // enough to guarantee that the whitespace cell isn't wider due to chance.
            // Experimentally, I calculated the ratio as 1.2478181, though this ratio may change
            // if the default font or margins change.
            double expectedRatioThreshold = 1.2f;
            double leadingWhitespaceRatio = ((double)leadingWhitespaceColWidth) / noWhitespaceColWidth;
            double trailingWhitespaceRatio = ((double)leadingWhitespaceColWidth) / noWhitespaceColWidth;

            assertGreaterThan("leading whitespace is longer than no whitespace", leadingWhitespaceRatio, expectedRatioThreshold);
            assertGreaterThan("trailing whitespace is longer than no whitespace", trailingWhitespaceRatio, expectedRatioThreshold);
            Assert.AreEqual(leadingWhitespaceColWidth, trailingWhitespaceColWidth,
                "cells with equal leading and trailing whitespace have equal width");

            wb.Close();
        }

        /**
         * Test if a > b. Fails if false.
         */
        private void assertGreaterThan(String message, double a, double b)
        {
            if (a <= b)
            {
                String msg = "Expected: " + a + " > " + b;
                Assert.Fail(message + ": " + msg);
            }
        }

        private double ComputeCellWidthManually(ICell cell0, IFont font)
        {
            double width;
            //FontRenderContext fontRenderContext = new FontRenderContext(null, true, true);
            IRichTextString rt = cell0.RichStringCellValue;
            String[] lines = rt.String.Split("\n".ToCharArray());
            Assert.AreEqual(1, lines.Length);
            String txt = lines[0] + "0";

            //AttributedString str = new AttributedString(txt);
            //copyAttributes(font, str, 0, txt.length());
            // TODO: support rich text fragments
            //if (rt.NumFormattingRuns > 0)
            //{
            //}

            //TextLayout layout = new TextLayout(str.getIterator(), fontRenderContext);
            //width = ((layout.getBounds().getWidth() / 1) / 8);
            Font wfont = SheetUtil.IFont2Font(font);
            width= (double)TextMeasurer.Measure(txt, new TextOptions(wfont)).Width;
            return width;
        }

        private double ComputeCellWidthFixed(IFont font, String txt)
        {
            double width;
            Font wfont = SheetUtil.IFont2Font(font);
            width = (double)TextMeasurer.Measure(txt, new TextOptions(wfont)).Width;
            return width;
        }
        
        //private static void copyAttributes(Font font, AttributedString str, int startIdx, int endIdx)
        //{
        //    str.addAttribute(TextAttribute.FAMILY, font.getFontName(), startIdx, endIdx);
        //    str.addAttribute(TextAttribute.SIZE, (float)font.getFontHeightInPoints());
        //    if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD) str.addAttribute(TextAttribute.WEIGHT, TextAttribute.WEIGHT_BOLD, startIdx, endIdx);
        //    if (font.getItalic()) str.addAttribute(TextAttribute.POSTURE, TextAttribute.POSTURE_OBLIQUE, startIdx, endIdx);
        //    if (font.getUnderline() == Font.U_SINGLE) str.addAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON, startIdx, endIdx);
        //}

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
            _testDataProvider.TrackAllColumnsForAutosizing(s);
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
                r1.GetCell(i).CellStyle = (/*setter*/iPercent);
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
            Assert.AreEqual(FormulaError.VALUE.Code, cf.ErrorCellValue);

            cf.SetCellFormula("SEARCH(\"am\", B1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(19, (int)cf.NumericCellValue);

            cf.SetCellFormula("SEARCH(\"am\", C1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(FormulaError.VALUE.Code, cf.ErrorCellValue);

            cf.SetCellFormula("SEARCH(\"am\", D1)");
            cf = EvaluateCell(wb, cf);
            Assert.AreEqual(FormulaError.VALUE.Code, cf.ErrorCellValue);


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
        [Test]
        [Ignore("Fix this to evaluate for XSSF, Fix this to work at all for HSSF")]
        public void Bug46670()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r1 = s.CreateRow(0);


            // References to try
            String ext = _testDataProvider.StandardFileNameExtension;
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
        [Test]
        public void Test56574OverwriteExistingRow()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            { // create the Formula-Cell
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.SetCellFormula("A2");
            }

            { // check that it is there now
                IRow row = sheet.GetRow(0);

                /* CTCell[] cArray = ((XSSFRow)row).getCTRow().getCArray();
                 assertEquals(1, cArray.length);*/

                ICell cell = row.GetCell(0);
                Assert.AreEqual(CellType.Formula, cell.CellType);
            }

            { // overwrite the row
                IRow row = sheet.CreateRow(0);
                Assert.IsNotNull(row);
            }

            { // creating a row in place of another should remove the existing data,
              // check that the cell is gone now
                IRow row = sheet.GetRow(0);

                /*CTCell[] cArray = ((XSSFRow)row).getCTRow().getCArray();
                assertEquals(0, cArray.length);*/

                ICell cell = row.GetCell(0);
                Assert.IsNull(cell);
            }

            // the calculation chain in XSSF is empty in a newly created workbook, so we cannot check if it is correctly updated
            /*assertNull(((XSSFWorkbook)wb).getCalculationChain());
            assertNotNull(((XSSFWorkbook)wb).getCalculationChain().getCTCalcChain());
            assertNotNull(((XSSFWorkbook)wb).getCalculationChain().getCTCalcChain().getCArray());
            assertEquals(0, ((XSSFWorkbook)wb).getCalculationChain().getCTCalcChain().getCArray().length);*/

            wb.Close();
        }

        /**
         * With HSSF, if you create a font, don't change it, and
         *  create a 2nd, you really do get two fonts that you
         *  can alter as and when you want.
         * With XSSF, that wasn't the case, but this verfies
         *  that it now is again
         */
        [Test]
        public void Bug48718()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            int startingFonts = wb is HSSFWorkbook ? 4 : 1;

            Assert.AreEqual(startingFonts, wb.NumberOfFonts);

            // Get a font, and slightly change it
            IFont a = wb.CreateFont();
            Assert.AreEqual(startingFonts + 1, wb.NumberOfFonts);
            a.FontHeightInPoints = ((short)23);
            Assert.AreEqual(startingFonts + 1, wb.NumberOfFonts);

            // Get two more, unchanged
            /*Font b =*/
            wb.CreateFont();
            Assert.AreEqual(startingFonts + 2, wb.NumberOfFonts);
            /*Font c =*/
            wb.CreateFont();
            Assert.AreEqual(startingFonts + 3, wb.NumberOfFonts);

            wb.Close();
        }

        [Test]
        public void Bug57430()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet1");

            IName name1 = wb.CreateName();
            name1.NameName = ("FMLA");
            name1.RefersToFormula = ("Sheet1!$B$3");
            wb.Close();
        }

        [Test]
        public void Bug56981()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICellStyle vertTop = wb.CreateCellStyle();
            vertTop.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Top;
            ICellStyle vertBottom = wb.CreateCellStyle();
            vertBottom.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Bottom;
            ISheet sheet = wb.CreateSheet("Sheet 1");
            IRow row = sheet.CreateRow(0);
            ICell top = row.CreateCell(0);
            ICell bottom = row.CreateCell(1);
            top.SetCellValue("Top");
            top.CellStyle = (vertTop); // comment this out to get all bottom-aligned
                                       // cells
            bottom.SetCellValue("Bottom");
            bottom.CellStyle = (vertBottom);
            row.HeightInPoints = (85.75f); // make it obvious

            /*FileOutputStream out = new FileOutputStream("c:\\temp\\56981.xlsx");
            try {
                wb.write(out);
            } finally {
                out.close();
            }*/

            wb.Close();
        }

        [Test]
        public void Test57973()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();

            ICreationHelper factory = wb.GetCreationHelper();

            ISheet sheet = wb.CreateSheet();
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = factory.CreateClientAnchor();

            ICell cell0 = sheet.CreateRow(0).CreateCell(0);
            cell0.SetCellValue("Cell0");

            IComment comment0 = drawing.CreateCellComment(anchor);
            IRichTextString str0 = factory.CreateRichTextString("Hello, World1!");
            comment0.String = (str0);
            comment0.Author = ("Apache POI");
            cell0.CellComment = (comment0);

            anchor = factory.CreateClientAnchor();
            anchor.Col1 = (1);
            anchor.Col2 = (1);
            anchor.Row1 = (1);
            anchor.Row2 = (1);
            ICell cell1 = sheet.CreateRow(3).CreateCell(5);
            cell1.SetCellValue("F4");
            IComment comment1 = drawing.CreateCellComment(anchor);
            IRichTextString str1 = factory.CreateRichTextString("Hello, World2!");
            comment1.String = (str1);
            comment1.Author = ("Apache POI");
            cell1.CellComment = (comment1);

            ICell cell2 = sheet.CreateRow(2).CreateCell(2);
            cell2.SetCellValue("C3");

            anchor = factory.CreateClientAnchor();
            anchor.Col1 = (2);
            anchor.Col2 = (2);
            anchor.Row1 = (2);
            anchor.Row2 = (2);

            IComment comment2 = drawing.CreateCellComment(anchor);
            IRichTextString str2 = factory.CreateRichTextString("XSSF can set cell comments");
            //apply custom font to the text in the comment
            IFont font = wb.CreateFont();
            font.FontName = ("Arial");
            font.FontHeightInPoints = ((short)14);
            font.IsBold=true;
            font.Color = (IndexedColors.Red.Index);
            str2.ApplyFont(font);

            comment2.String = (str2);
            comment2.Author = ("Apache POI");
            comment2.Column = (2);
            comment2.Row = (2);

            wb.Close();
        }

        /**
         * Ensures that XSSF and HSSF agree with each other,
         *  and with the docs on when fetching the wrong
         *  kind of value from a Formula cell
         */
        [Test]
        public virtual void Bug47815()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);

            // Setup
            ICell cn = r.CreateCell(0, CellType.Numeric);
            cn.SetCellValue(1.2);
            ICell cs = r.CreateCell(1, CellType.String);
            cs.SetCellValue("Testing");

            ICell cfn = r.CreateCell(2, CellType.Formula);
            cfn.SetCellFormula("A1");
            ICell cfs = r.CreateCell(3, CellType.Formula);
            cfs.SetCellFormula("B1");

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual(CellType.Numeric, fe.Evaluate(cfn).CellType);
            Assert.AreEqual(CellType.String, fe.Evaluate(cfs).CellType);
            fe.EvaluateFormulaCell(cfn);
            fe.EvaluateFormulaCell(cfs);

            // Now test
            Assert.AreEqual(CellType.Numeric, cn.CellType);
            Assert.AreEqual(CellType.String, cs.CellType);
            Assert.AreEqual(CellType.Formula, cfn.CellType);
            Assert.AreEqual(CellType.Numeric, cfn.CachedFormulaResultType);
            Assert.AreEqual(CellType.Formula, cfs.CellType);
            Assert.AreEqual(CellType.String, cfs.CachedFormulaResultType);

            // Different ways of retrieving
            Assert.AreEqual(1.2, cn.NumericCellValue, 0);
            try
            {
                var tmp = cn.RichStringCellValue;
                Assert.Fail();
            }
            catch (InvalidOperationException) { }

            Assert.AreEqual("Testing", cs.StringCellValue);
            try
            {
                var tmp = cs.NumericCellValue;
                Assert.Fail();
            }
            catch (InvalidOperationException) { }

            Assert.AreEqual(1.2, cfn.NumericCellValue, 0);
            try
            {
                var tmp = cfn.RichStringCellValue;
                Assert.Fail();
            }
            catch (InvalidOperationException) { }

            Assert.AreEqual("Testing", cfs.StringCellValue);
            try
            {
                var tmp = cfs.NumericCellValue;
                Assert.Fail();
            }
            catch (InvalidOperationException) { }

            wb.Close();
        }
        [Test]
        public virtual void Test58113()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            // verify that null-values can be set, this was possible up to 3.11, but broken in 3.12 
            cell.SetCellValue((String)null);
            String value = cell.StringCellValue;
            Assert.IsTrue(value == null || value.Length == 0, "HSSF will currently return empty string, XSSF/SXSSF will return null, but had: " + value);

            cell = row.CreateCell(1);
            // also verify that setting formulas to null works  
            cell.SetCellType(CellType.Formula);
            cell.SetCellValue((String)null);

            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
            value = cell.StringCellValue;
            Assert.IsTrue(value == null || value.Length == 0, "HSSF will currently return empty string, XSSF/SXSSF will return null, but had: " + value);

            // set some value
            cell.SetCellType(CellType.String);
            cell.SetCellValue("somevalue");
            value = cell.StringCellValue;
            Assert.IsTrue(value.Equals("somevalue"), "can set value afterwards: " + value);
            // verify that the null-value is actually set even if there was some value in the cell before  
            cell.SetCellValue((String)null);
            value = cell.StringCellValue;
            Assert.IsTrue(value == null || value.Length == 0, "HSSF will currently return empty string, XSSF/SXSSF will return null, but had: " + value);
        }

        /**
         * Formulas with Nested Ifs, or If with text functions like
         *  Mid in it, can give #VALUE in Excel
         */
        [Test]
        public void Bug55747()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IFormulaEvaluator ev = wb.GetCreationHelper().CreateFormulaEvaluator();
            ISheet s = wb.CreateSheet();

            IRow row = s.CreateRow(0);
            row.CreateCell(0).SetCellValue("abc");
            row.CreateCell(1).SetCellValue("");
            row.CreateCell(2).SetCellValue(3);
            ICell cell = row.CreateCell(5);
            cell.SetCellFormula("IF(A1<>\"\",MID(A1,1,2),\" \")");
            ev.EvaluateAll();
            Assert.AreEqual("ab", cell.StringCellValue);

            cell = row.CreateCell(6);
            cell.SetCellFormula("IF(B1<>\"\",MID(A1,1,2),\"empty\")");
            ev.EvaluateAll();
            Assert.AreEqual("empty", cell.StringCellValue);

            cell = row.CreateCell(7);
            cell.SetCellFormula("IF(A1<>\"\",IF(C1<>\"\",MID(A1,1,2),\"c1\"),\"c2\")");
            ev.EvaluateAll();
            Assert.AreEqual("ab", cell.StringCellValue);


            // Write it back out, and re-read
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            ev = wb.GetCreationHelper().CreateFormulaEvaluator();
            s = wb.GetSheetAt(0);
            row = s.GetRow(0);

            // Check read ok, and re-evaluate fine
            cell = row.GetCell(5);
            Assert.AreEqual("ab", cell.StringCellValue);
            ev.EvaluateFormulaCell(cell);
            Assert.AreEqual("ab", cell.StringCellValue);

            cell = row.GetCell(6);
            Assert.AreEqual("empty", cell.StringCellValue);
            ev.EvaluateFormulaCell(cell);
            Assert.AreEqual("empty", cell.StringCellValue);

            cell = row.GetCell(7);
            Assert.AreEqual("ab", cell.StringCellValue);
            ev.EvaluateFormulaCell(cell);
            Assert.AreEqual("ab", cell.StringCellValue);

        }

        [Test]
        public void Bug58260()
        {
            //Create workbook and worksheet
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            //ISheet worksheet = wb.CreateSheet("sample");
            //Loop through and add all values from array list
            // use a fixed seed to always produce the same file which makes comparing stuff easier
            //Random rnd = new Random(4352345);
            int maxStyles = (wb is HSSFWorkbook) ? 4009 : 64000;
            for (int i = 0; i < maxStyles; i++)
            {
                //Create new row
                //IRow row = worksheet.CreateRow(i);
                //Create cell style
                ICellStyle style;
                try
                {
                    style = wb.CreateCellStyle();
                }
                catch (InvalidOperationException e)
                {
                    throw new InvalidOperationException("Failed for row " + i, e);
                }
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
                if ((wb is HSSFWorkbook))
                {
                    // there are some predefined styles
                    Assert.AreEqual(i + 21, style.Index);
                }
                else
                {
                    // getIndex() returns short, which is not sufficient for > 32767
                    // we should really change the API to be "int" for getIndex() but
                    // that needs API changes
                    Assert.AreEqual(i + 1, style.Index & 0xffff);
                }
                //Create cell
                //ICell cell = row.CreateCell(0);
                //Set cell style
                //cell.CellStyle = (style);
                //Set cell value
                //cell.SetCellValue("r" + rnd.Next());
            }
            // should Assert.Fail if we try to add more now
            try
            {
                wb.CreateCellStyle();
                Assert.Fail("Should Assert.Fail after " + maxStyles + " styles, but did not Assert.Fail");
            }
            catch (InvalidOperationException)
            {
                // expected here
            }

            /*//add column width for appearance sake
            worksheet.setColumnWidth(0, 5000);

            // Write the output to a file       
            Console.WriteLine("Writing...");
            OutputStream fileOut = new FileOutputStream("C:\\temp\\58260." + _testDataProvider.StandardFileNameExtension);

            // the resulting file can be compressed nicely, so we need to disable the zip bomb detection here
            double before = ZipSecureFile.MinInflateRatio;
            try {
                ZipSecureFile.setMinInflateRatio(0.00001);
                wb.write(fileOut);
            } finally { 
                fileOut.close();
                ZipSecureFile.setMinInflateRatio(before);
            }*/

            wb.Close();
        }

        [Test]
        public void Test50319()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            sheet.CreateRow(0);
            sheet.GroupRow(0, 0);
            sheet.SetRowGroupCollapsed(0, true);

            sheet.GroupColumn(0, 0);
            sheet.SetColumnGroupCollapsed(0, true);

            wb.Close();
        }

        [Ignore("by poi")]
        [Test]
        public void test58648()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            cell.CellFormula = ("((1 + 1) )");
            // Assert.Fails with
            // org.apache.poi.ss.formula.FormulaParseException: Parse error near char ... ')'
            // in specified formula '((1 + 1) )'. Expected cell ref or constant literal
            wb.Close();
        }

        /**
        * If someone sets a null string as a cell value, treat
        *  it as an empty cell, and avoid a NPE on auto-sizing
        */
        [Test]
        public void Test57034()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            ICell cell = s.CreateRow(0).CreateCell(0);
            cell.SetCellValue((String)null);
            Assert.AreEqual(CellType.Blank, cell.CellType);

            _testDataProvider.TrackAllColumnsForAutosizing(s);

            s.AutoSizeColumn(0);
            Assert.AreEqual(2048, s.GetColumnWidth(0));
            s.AutoSizeColumn(0, true);
            Assert.AreEqual(2048, s.GetColumnWidth(0));
            wb.Close();
        }

        [Test]
        public void Test52684()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(12312345123L);
            IDataFormat format = wb.CreateDataFormat();
            ICellStyle style = wb.CreateCellStyle();
            style.DataFormat = (format.GetFormat("000-00000-000"));
            cell.CellStyle = (style);
            Assert.AreEqual("000-00000-000",
                    cell.CellStyle.GetDataFormatString());
            Assert.AreEqual(164, cell.CellStyle.DataFormat);
            DataFormatter formatter = new DataFormatter();
            Assert.AreEqual("12-312-345-123", formatter.FormatCellValue(cell));
            wb.Close();
        }


        [Test]
        public void Test58896()
        {
            int nrows = 160;
            int ncols = 139;
            TextWriter out1 = Console.Out;

            // Create a workbook
            IWorkbook wb = _testDataProvider.CreateWorkbook(nrows + 1);
            ISheet sh = wb.CreateSheet();
            out1.WriteLine(wb.GetType().Name + " column autosizing timing...");

            long t0 = TimeUtil.CurrentMillis();
            _testDataProvider.TrackAllColumnsForAutosizing(sh);
            for (int r = 0; r < nrows; r++)
            {
                IRow row = sh.CreateRow(r);
                for (int c = 0; c < ncols; c++)
                {
                    ICell cell = row.CreateCell(c);
                    cell.SetCellValue("Cell[r=" + r + ",c=" + c + "]");
                }
            }
            double populateSheetTime = delta(t0);
            double populateSheetTimePerCell_ns = (1000000 * populateSheetTime / (nrows * ncols));
            out1.WriteLine("Populate sheet time: " + populateSheetTime + " ms (" + populateSheetTimePerCell_ns + " ns/cell)");

            out1.WriteLine("\nAutosizing...");
            long t1 = TimeUtil.CurrentMillis();
            for (int c = 0; c < ncols; c++)
            {
                long t2 = TimeUtil.CurrentMillis();
                sh.AutoSizeColumn(c);
                out1.WriteLine("Column " + c + " took " + delta(t2) + " ms");
            }
            double autoSizeColumnsTime = delta(t1);
            double autoSizeColumnsTimePerColumn = autoSizeColumnsTime / ncols;
            double bestFitWidthTimePerCell_ns = 1000000 * autoSizeColumnsTime / (ncols * nrows);

            out1.WriteLine("Auto sizing columns took a total of " + autoSizeColumnsTime + " ms (" + autoSizeColumnsTimePerColumn + " ms per column)");
            out1.WriteLine("Best fit width time per cell: " + bestFitWidthTimePerCell_ns + " ns");

            double totalTime_s = (populateSheetTime + autoSizeColumnsTime) / 1000;
            out1.WriteLine("Total time: " + totalTime_s + " s");

            wb.Close();

            //if (bestFitWidthTimePerCell_ns > 50000) {
            //    Assert.Fail("Best fit width time per cell exceeded 50000 ns: " + bestFitWidthTimePerCell_ns + " ns");
            //}

            //if (totalTime_s > 10)
            //{
            //    Assert.Fail("Total time exceeded 10 seconds: " + totalTime_s + " s");
            //}
        }

        protected double delta(long startTimeMillis)
        {
            return TimeUtil.CurrentMillis() - startTimeMillis;
        }

        [Ignore("bug 59393")]
        [Test]
        public void Bug59393_commentsCanHaveSameAnchor()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            ISheet sheet = wb.CreateSheet();

            ICreationHelper helper = wb.GetCreationHelper();
            IClientAnchor anchor = helper.CreateClientAnchor();
            IDrawing drawing = sheet.CreateDrawingPatriarch();

            IRow row = sheet.CreateRow(0);

            ICell cell1 = row.CreateCell(0);
            ICell cell2 = row.CreateCell(1);
            ICell cell3 = row.CreateCell(2);
            IComment comment1 = drawing.CreateCellComment(anchor);
            IRichTextString richTextString1 = helper.CreateRichTextString("comment1");
            comment1.String = richTextString1;
            cell1.CellComment = comment1;

            // Assert.Fails with IllegalArgumentException("Multiple cell comments in one cell are not allowed, cell: A1")
            // because createCellComment tries to create a cell at A1
            // (from CellAddress(anchor.Row1, anchor.Cell1)),
            // but cell A1 already has a comment (comment1).
            // Need to atomically create a comment and attach it to a cell.
            // Current workaround: change anchor between each usage
            // anchor.Col1=1;
            IComment comment2 = drawing.CreateCellComment(anchor);
            IRichTextString richTextString2 = helper.CreateRichTextString("comment2");
            comment2.String = richTextString2;
            cell2.CellComment = comment2;
            // anchor.Col1=2;
            IComment comment3 = drawing.CreateCellComment(anchor);
            IRichTextString richTextString3 = helper.CreateRichTextString("comment3");
            comment3.String = richTextString3;
            cell3.CellComment = comment3;

            wb.Close();
        }

        [Test]
        public virtual void Bug57798()
        {
            String fileName = "57798." + _testDataProvider.StandardFileNameExtension;
            IWorkbook workbook = _testDataProvider.OpenSampleWorkbook(fileName);
            ISheet sheet = workbook.GetSheet("Sheet1");
            // *******************************
            // First cell of array formula, OK
            int rowId = 0;
            int cellId = 1;
            Console.WriteLine("Reading row " + rowId + ", col " + cellId);
            IRow row = sheet.GetRow(rowId);
            ICell cell = row.GetCell(cellId);
            Console.WriteLine("Formula:" + cell.CellFormula);
            if (CellType.Formula == cell.CellType)
            {
                int formulaResultType = (int)cell.CachedFormulaResultType;
                Console.WriteLine("Formual Result Type:" + formulaResultType);
            }
            // *******************************
            // Second cell of array formula, NOT OK for xlsx files
            rowId = 1;
            cellId = 1;
            Console.WriteLine("Reading row " + rowId + ", col " + cellId);
            row = sheet.GetRow(rowId);
            cell = row.GetCell(cellId);
            Console.WriteLine("Formula:" + cell.CellFormula);
            if (CellType.Formula == cell.CellType)
            {
                int formulaResultType = (int)cell.CachedFormulaResultType;
                Console.WriteLine("Formual Result Type:" + formulaResultType);
            }
            workbook.Close();
        }


        [Ignore("")]
        [Test]
        public void Test57929()
        {
            // Create a workbook with print areas on 2 sheets
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet0");
            wb.CreateSheet("Sheet1");
            wb.SetPrintArea(0, "$A$1:$C$6");
            wb.SetPrintArea(1, "$B$1:$C$5");

            // Verify the print areas were set correctly
            Assert.AreEqual("Sheet0!$A$1:$C$6", wb.GetPrintArea(0));
            Assert.AreEqual("Sheet1!$B$1:$C$5", wb.GetPrintArea(1));

            // Remove the print area on Sheet0 and change the print area on Sheet1
            wb.RemovePrintArea(0);
            wb.SetPrintArea(1, "$A$1:$A$1");

            // Verify that the changes were made
            Assert.IsNull(wb.GetPrintArea(0), "Sheet0 before write");
            Assert.AreEqual("Sheet1!$A$1:$A$1", wb.GetPrintArea(1), "Sheet1 before write");

            // Verify that the changes are non-volatile
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb);
            wb.Close();

            Assert.IsNull(wb2.GetPrintArea(0), "Sheet0 after write"); // CURRENTLY FAILS with "Sheet0!$A$1:$C$6"
            Assert.AreEqual("Sheet1!$A$1:$A$1", wb2.GetPrintArea(1), "Sheet1 after write");
        }
        [Test]
        public void test55384()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sh = wb.CreateSheet();
                for (int rownum = 0; rownum < 10; rownum++)
                {
                    IRow row1 = sh.CreateRow(rownum);
                    for (int cellnum = 0; cellnum < 3; cellnum++)
                    {
                        ICell cell = row1.CreateCell(cellnum);
                        cell.SetCellValue(rownum + cellnum);
                    }
                }
                IRow row = sh.CreateRow(10);
                // setting no precalculated value works just fine.
                ICell cell1 = row.CreateCell(0);
                cell1.SetCellFormula("SUM(A1:A10)");
                // but setting a precalculated STRING value Assert.Fails totally in1 SXSSF
                ICell cell2 = row.CreateCell(1);
                cell2.SetCellFormula("SUM(B1:B10)");
                cell2.SetCellValue("55");
                // setting a precalculated int value works as expected
                ICell cell3 = row.CreateCell(2);
                cell3.SetCellFormula("SUM(C1:C10)");
                cell3.SetCellValue(65);
                Assert.AreEqual(CellType.Formula, cell1.CellType);
                Assert.AreEqual(CellType.Formula, cell2.CellType);
                Assert.AreEqual(CellType.Formula, cell3.CellType);
                Assert.AreEqual("SUM(A1:A10)", cell1.CellFormula);
                Assert.AreEqual("SUM(B1:B10)", cell2.CellFormula);
                Assert.AreEqual("SUM(C1:C10)", cell3.CellFormula);
                /*String name = wb.GetClass().getCanonicalName();
                String ext = (wb is HSSFWorkbook) ? ".xls" : ".xlsx";
                OutputStream output = new FileOutputStream("/tmp" + name + ext);
                try {
                    wb.Write(output);
                } finally {
                    output.Close();
                }*/
                IWorkbook wbBack = _testDataProvider.WriteOutAndReadBack(wb);
                checkFormulaPreevaluatedString(wbBack);
                wbBack.Close();
            }
            finally
            {
                wb.Close();
            }
        }
        private void checkFormulaPreevaluatedString(IWorkbook readFile)
        {
            ISheet sheet = readFile.GetSheetAt(0);
            IRow row = sheet.GetRow(sheet.LastRowNum);
            Assert.AreEqual(10, row.RowNum);
            foreach (ICell cell in row)
            {
                String cellValue = null;
                switch (cell.CellType)
                {
                    case CellType.String:
                        cellValue = cell.RichStringCellValue.String;
                        break;
                    case CellType.Formula:
                        cellValue = cell.CellFormula;
                        break;
                }
                Assert.IsNotNull(cellValue);
                cellValue = string.IsNullOrEmpty(cellValue) ? null : cellValue;
                Assert.IsNotNull(cellValue);
            }
        }

    }
}
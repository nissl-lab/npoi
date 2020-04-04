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

namespace TestCases.XSSF.Model
{
    using System;
    using System.IO;
    using NUnit.Framework;
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using System.Collections.Generic;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.XSSF.Model;

    [TestFixture]
    public class TestThemesTable
    {
        private String testFileSimple = "Themes.xlsx";
        private String testFileComplex = "Themes2.xlsx";
        // TODO .xls version available too, add HSSF support then check 

        // What colours they should show up as
        private static String[] rgbExpected = {
            "ffffff", // Lt1
            "000000", // Dk1
            "eeece1", // Lt2
            "1f497d", // DK2
            "4f81bd", // Accent1
            "c0504d", // Accent2
            "9bbb59", // Accent3
            "8064a2", // Accent4
            "4bacc6", // Accent5
            "f79646", // Accent6
            "0000ff", // Hlink
            "800080"  // FolHlink
        };
        [Test]
        public void TestThemesTableColors()
        {
            // Load our two test workbooks
            XSSFWorkbook simple = XSSFTestDataSamples.OpenSampleWorkbook(testFileSimple);
            XSSFWorkbook complex = XSSFTestDataSamples.OpenSampleWorkbook(testFileComplex);
            // Save and re-load them, to check for stability across that
            XSSFWorkbook simpleRS = XSSFTestDataSamples.WriteOutAndReadBack(simple) as XSSFWorkbook;
            XSSFWorkbook complexRS = XSSFTestDataSamples.WriteOutAndReadBack(complex) as XSSFWorkbook;
            // Fetch fresh copies to test with
            simple = XSSFTestDataSamples.OpenSampleWorkbook(testFileSimple);
            complex = XSSFTestDataSamples.OpenSampleWorkbook(testFileComplex);
            // Files and descriptions
            Dictionary<String, XSSFWorkbook> workbooks = new Dictionary<String, XSSFWorkbook>();
            workbooks.Add(testFileSimple, simple);
            workbooks.Add("Re-Saved_" + testFileSimple, simpleRS);
            workbooks.Add(testFileComplex, complex);
            workbooks.Add("Re-Saved_" + testFileComplex, complexRS);

            // Sanity check
            //Assert.AreEqual(rgbExpected.Length, rgbExpected.Length);

            // For offline testing
            bool createFiles = false;

            // Check each workbook in turn, and verify that the colours
            //  for the theme-applied cells in Column A are correct
            foreach (String whatWorkbook in workbooks.Keys)
            {
                XSSFWorkbook workbook = workbooks[whatWorkbook];
                XSSFSheet sheet = workbook.GetSheetAt(0) as XSSFSheet;
                int startRN = 0;
                if (whatWorkbook.EndsWith(testFileComplex)) startRN++;

                for (int rn = startRN; rn < rgbExpected.Length + startRN; rn++)
                {
                    XSSFRow row = sheet.GetRow(rn) as XSSFRow;
                    Assert.IsNotNull(row, "Missing row " + rn + " in " + whatWorkbook);
                    String ref1 = (new CellReference(rn, 0)).FormatAsString();
                    XSSFCell cell = row.GetCell(0) as XSSFCell;
                    Assert.IsNotNull(cell,
                            "Missing cell " + ref1 +" in " + whatWorkbook);

                    int expectedThemeIdx = rn - startRN;
                    ThemeElement themeElem = ThemeElement.ById(expectedThemeIdx);
                    Assert.AreEqual(themeElem.name.ToLower(), cell.StringCellValue,
                            "Wrong theme at " + ref1 +" in " + whatWorkbook);

                    // Fonts are theme-based in their colours
                    XSSFFont font = (cell.CellStyle as XSSFCellStyle).GetFont();
                    CT_Color ctColor = font.GetCTFont().GetColorArray(0);
                    Assert.IsNotNull(ctColor);
                    Assert.AreEqual(true, ctColor.IsSetTheme());
                    Assert.AreEqual(themeElem.idx, ctColor.theme);

                    // Get the colour, via the theme
                    XSSFColor color = font.GetXSSFColor();

                    // Theme colours aren't tinted
                    Assert.AreEqual(color.HasTint, false);
                    // Check the RGB part (no tint)
                    Assert.AreEqual(rgbExpected[expectedThemeIdx], HexDump.EncodeHexString(color.RGB),
                            "Wrong theme colour " + themeElem.name + " on " + whatWorkbook);
                    
                    long themeIdx = font.GetCTFont().GetColorArray(0).theme;
                    Assert.AreEqual(expectedThemeIdx, themeIdx,
                            "Wrong theme index " + expectedThemeIdx + " on " + whatWorkbook
                            );

                    if (createFiles)
                    {
                        XSSFCellStyle cs = row.Sheet.Workbook.CreateCellStyle() as XSSFCellStyle;
                        cs.SetFillForegroundColor(color);
                        cs.FillPattern = FillPattern.SolidForeground;
                        row.CreateCell(1).CellStyle = (cs);
                    }
                }

                if (createFiles)
                {
                    FileStream fos = new FileStream("Generated_" + whatWorkbook, FileMode.Create, FileAccess.ReadWrite);
                    workbook.Write(fos);
                    fos.Close();
                }
            }
        }

        /**
         * Ensure that, for a file with themes, we can correctly
         *  read both the themed and non-themed colours back.
         * Column A = Theme Foreground
         * Column B = Theme Foreground
         * Column C = Explicit Colour Foreground
         * Column E = Explicit Colour Background, Black Foreground
         * Column G = Conditional Formatting Backgrounds
         * (Row 4 = White by Lt2)
         * 
         * Note - Grey Row has an odd way of doing the styling... 
         */
        [Test]
        public void ThemedAndNonThemedColours()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook(testFileComplex);
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFCellStyle style;
            XSSFColor color;
            XSSFCell cell;

            String[] names = { "White", "Black", "Grey", "Dark Blue", "Blue", "Red", "Green" };
            String[] explicitFHexes = { "FFFFFFFF", "FF000000", "FFC0C0C0", "FF002060",
                                    "FF0070C0", "FFFF0000", "FF00B050" };
            String[] explicitBHexes = { "FFFFFFFF", "FF000000", "FFC0C0C0", "FF002060",
                                    "FF0000FF", "FFFF0000", "FF00FF00" };
            Assert.AreEqual(7, names.Length);

            // Check the non-CF colours in Columns A, B, C and E
            for (int rn = 1; rn < 8; rn++)
            {
                int idx = rn - 1;
                XSSFRow row = sheet.GetRow(rn) as XSSFRow;
                Assert.IsNotNull(row, "Missing row " + rn);

                // Theme cells come first
                XSSFCell themeCell = row.GetCell(0) as XSSFCell;
                ThemeElement themeElem = ThemeElement.ById(idx);
                assertCellContents(themeElem.name, themeCell);
                // Sanity check names
                assertCellContents(names[idx], row.GetCell(1));
                assertCellContents(names[idx], row.GetCell(2));
                assertCellContents(names[idx], row.GetCell(4));

                // Check the colours

                //  A: Theme Based, Foreground
                style = themeCell.CellStyle as XSSFCellStyle;
                color = style.GetFont().GetXSSFColor();
                Assert.AreEqual(true, color.IsThemed);
                Assert.AreEqual(idx, color.Theme);
                Assert.AreEqual(rgbExpected[idx], HexDump.EncodeHexString(color.RGB));
                //  B: Theme Based, Foreground
                cell = row.GetCell(1) as XSSFCell;
                style = cell.CellStyle as XSSFCellStyle;
                color = style.GetFont().GetXSSFColor();
                Assert.AreEqual(true, color.IsThemed);
                // TODO Fix the grey theme color in Column B
                if (idx != 2)
                {
                    Assert.AreEqual(true, color.IsThemed);
                    Assert.AreEqual(idx, color.Theme);
                    Assert.AreEqual(rgbExpected[idx], HexDump.EncodeHexString(color.RGB));
                }
                else
                {
                    Assert.AreEqual(1, color.Theme);
                    Assert.AreEqual(0.50, color.Tint, 0.001);
                }

                //  C: Explicit, Foreground
                cell = row.GetCell(2) as XSSFCell;
                style = cell.CellStyle as XSSFCellStyle;
                color = style.GetFont().GetXSSFColor();
                Assert.AreEqual(false, color.IsThemed);
                Assert.AreEqual(explicitFHexes[idx], color.ARGBHex);

                // E: Explicit Background, Foreground all Black
                cell = row.GetCell(4) as XSSFCell;
                style = cell.CellStyle as XSSFCellStyle;

                color = style.GetFont().GetXSSFColor();
                Assert.AreEqual(true, color.IsThemed);
                Assert.AreEqual("FF000000", color.ARGBHex);

                color = style.FillForegroundXSSFColor;
                Assert.AreEqual(false, color.IsThemed);
                Assert.AreEqual(explicitBHexes[idx], color.ARGBHex);
                color = style.FillBackgroundColorColor as XSSFColor;
                Assert.AreEqual(false, color.IsThemed);
                Assert.AreEqual(null, color.ARGBHex);
            }

            // Check the CF colours
            // TODO
        }
        private static void assertCellContents(String expected, ICell cell)
        {
            Assert.IsNotNull(cell);
            Assert.AreEqual(expected.ToLower(), cell.StringCellValue.ToLower());
        }


        [Test]
        public void TestAddNew()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet s = wb.CreateSheet() as XSSFSheet;
            Assert.AreEqual(null, wb.GetTheme());

            StylesTable styles = wb.GetStylesSource();
            Assert.AreEqual(null, styles.GetTheme());

            styles.EnsureThemesTable();

            Assert.IsNotNull(styles.GetTheme());
            Assert.IsNotNull(wb.GetTheme());

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            styles = wb.GetStylesSource();
            Assert.IsNotNull(styles.GetTheme());
            Assert.IsNotNull(wb.GetTheme());
        }
    }
}
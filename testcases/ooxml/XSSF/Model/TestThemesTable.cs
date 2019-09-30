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

namespace NPOI.XSSF.Model
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

    [TestFixture]
    public class TestThemesTable
    {
        private String testFileSimple = "Themes.xlsx";
        private String testFileComplex = "Themes2.xlsx";
        // TODO .xls version available too, add HSSF support then check 

        // What our theme names are
        private static String[] themeEntries = {
            "lt1",
            "dk1",
            "lt2",
            "dk2",
            "accent1",
            "accent2",
            "accent3",
            "accent4",
            "accent5",
            "accent6",
            "hlink",
            "folhlink"
        };
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
            // TODO Fix these to work!
            //        workbooks.put(testFileComplex, complex);
            //        workbooks.put("Re-Saved_" + testFileComplex, complexRS);

            // Sanity check
            Assert.AreEqual(themeEntries.Length, rgbExpected.Length);

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

                for (int rn = startRN; rn < themeEntries.Length + startRN; rn++)
                {
                    XSSFRow row = sheet.GetRow(rn) as XSSFRow;
                    Assert.IsNotNull(row, "Missing row " + rn + " in " + whatWorkbook);
                    String ref1 = (new CellReference(rn, 0)).FormatAsString();
                    XSSFCell cell = row.GetCell(0) as XSSFCell;
                    Assert.IsNotNull(cell,
                            "Missing cell " + ref1 +" in " + whatWorkbook);
                    Assert.AreEqual(
                            "Wrong theme at " + ref1 +" in " + whatWorkbook,
                            themeEntries[rn], cell.StringCellValue);
                    XSSFFont font = (cell.CellStyle as XSSFCellStyle).GetFont();
                    XSSFColor color = font.GetXSSFColor();

                    // Theme colours aren't tinted
                    Assert.AreEqual(color.HasTint, false);
                    // Check the RGB part (no tint)
                    Assert.AreEqual(rgbExpected[rn], HexDump.ToHex(color.RGB),
                            "Wrong theme colour " + themeEntries[rn] + " on " + whatWorkbook);
                    // Check the Theme ID
                    int expectedThemeIdx = rn - startRN;
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
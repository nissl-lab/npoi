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

using NPOI;
using NPOI.HSSF.Record.Aggregates;

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using System.Text;
    using System.Collections;
    using System.Linq;
    using NUnit.Framework;

    using TestCases.HSSF;

    using NPOI.HSSF.UserModel;
    //using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.SS.Formula;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Model;
    using System.Collections.Generic;
    using NPOI.SS.Formula.PTG;
    using NPOI.POIFS.FileSystem;
    using NPOI.HSSF.Extractor;
    using NPOI.HSSF.Record.Crypto;
    using NPOI.HSSF;
    using System.Net;
    using SixLabors.ImageSharp;

    /**
     * Testcases for bugs entered in bugzilla
     * the Test name contains the bugzilla bug id
     * @author Avik Sengupta
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestBugs
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }
        [TearDown]
        public void ResetPassword()
        {
            Biff8EncryptionKey.CurrentUserPassword = (null);
        }

        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        private static HSSFWorkbook WriteOutAndReadBack(HSSFWorkbook original)
        {
            return HSSFTestDataSamples.WriteOutAndReadBack(original);
        }

        private static void WriteTestOutputFileForViewing(HSSFWorkbook wb, String simpleFileName)
        {
            if (true)
            { // set to false to output Test files
                return;
            }
#if !HIDE_UNREACHABLE_CODE
            string file = TempFile.GetTempFilePath(simpleFileName + "#", ".xls");
            FileStream out1 = new FileStream(file, FileMode.Create);
            wb.Write(out1);
            out1.Close();

            if (!File.Exists(file))
            {
                throw new Exception("File was not written");
            }
            Console.WriteLine("Open file '" + Path.GetFullPath(file) + "' in Excel");
#endif
        }

        /** Test reading AND writing a complicated workbook
         *Test Opening resulting sheet in excel*/
        [Test]
        public void Test15228()
        {
            HSSFWorkbook wb = OpenSample("15228.xls");
            ISheet s = wb.GetSheetAt(0);
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);
            c.SetCellValue(10);
            WriteTestOutputFileForViewing(wb, "Test15228");
        }
        [Test]
        public void Test13796()
        {
            HSSFWorkbook wb = OpenSample("13796.xls");
            ISheet s = wb.GetSheetAt(0);
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);
            c.SetCellValue(10);
            WriteOutAndReadBack(wb);
        }
        /**Test writing a hyperlink
         * Open resulting sheet in Excel and Check that A1 contains a hyperlink*/
        [Test]
        [Ignore("not found in poi")]
        public void Test23094()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            r.CreateCell(0).CellFormula = ("HYPERLINK( \"http://jakarta.apache.org\", \"Jakarta\" )");

            WriteTestOutputFileForViewing(wb, "Test23094");
        }

        /** Test reading of a formula with a name and a cell ref in one
         **/
        [Test]
        public void Test14460()
        {
            HSSFWorkbook wb = OpenSample("14460.xls");
            wb.GetSheetAt(0);
        }
        [Test]
        public void Test14330()
        {
            HSSFWorkbook wb = OpenSample("14330-1.xls");
            wb.GetSheetAt(0);

            wb = OpenSample("14330-2.xls");
            wb.GetSheetAt(0);
        }

        private static void setCellText(ICell cell, String text)
        {
            cell.SetCellValue(new HSSFRichTextString(text));
        }

        /** Test rewriting a file with large number of unique strings
         *Open resulting file in Excel to Check results!*/
        [Test]
        public void Test15375()
        {
            HSSFWorkbook wb = OpenSample("15375.xls");
            ISheet sheet = wb.GetSheetAt(0);

            IRow row = sheet.GetRow(5);
            ICell cell = row.GetCell(3);
            if (cell == null)
                cell = row.CreateCell(3);

            // Write Test
            cell.SetCellType(CellType.String);
            setCellText(cell, "a Test");

            // change existing numeric cell value

            IRow oRow = sheet.GetRow(14);
            ICell oCell = oRow.GetCell(4);
            oCell.SetCellValue(75);
            oCell = oRow.GetCell(5);
            setCellText(oCell, "0.3");

            WriteTestOutputFileForViewing(wb, "Test15375");
        }

        /** Test writing a file with large number of unique strings
         *Open resulting file in Excel to Check results!*/
        [Test]
        public void Test15375_2()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            String tmp1 = null;
            String tmp2 = null;
            String tmp3 = null;

            for (int i = 0; i < 6000; i++)
            {
                tmp1 = "Test1" + i;
                tmp2 = "Test2" + i;
                tmp3 = "Test3" + i;

                IRow row = sheet.CreateRow(i);

                ICell cell = row.CreateCell(0);
                setCellText(cell, tmp1);
                cell = row.CreateCell(1);
                setCellText(cell, tmp2);
                cell = row.CreateCell(2);
                setCellText(cell, tmp3);
            }
            WriteTestOutputFileForViewing(wb, "Test15375-2");
        }
        /** another Test for the number of unique strings issue
         *Test Opening the resulting file in Excel*/
        [Test]
        [Ignore("this test was not found in poi 3.8beta4")]
        public void Test22568()
        {
            int r = 2000; int c = 3;

            HSSFWorkbook wb = new HSSFWorkbook();
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
                setCellText(cell, "Col " + (j + 1));
            }

            for (int i = 1; i < rw_cnt; i++)
            {
                rw = sheet.CreateRow(i);
                for (int j = 0; j < col_cnt; j++)
                {
                    ICell cell = rw.CreateCell(j);
                    setCellText(cell, "Row:" + (i + 1) + ",Column:" + (j + 1));
                }
            }

            sheet.DefaultColumnWidth = 18;

            WriteTestOutputFileForViewing(wb, "Test22568");
        }

        /**Double byte strings*/
        [Test]
        public void Test15556()
        {

            HSSFWorkbook wb = OpenSample("15556.xls");
            ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(45);
            Assert.IsNotNull(row, "Read row fine!");
        }
        /**Double byte strings */
        [Test]
        public void Test22742()
        {
            OpenSample("22742.xls");
        }
        /**Double byte strings */
        [Test]
        public void Test12561_1()
        {
            OpenSample("12561-1.xls");
        }
        /** Double byte strings */
        [Test]
        public void Test12561_2()
        {
            OpenSample("12561-2.xls");
        }
        /** Double byte strings
         File supplied by jubeson*/
        [Test]
        public void Test12843_1()
        {
            OpenSample("12843-1.xls");
        }

        /** Double byte strings
         File supplied by Paul Chung*/
        [Test]
        public void Test12843_2()
        {
            OpenSample("12843-2.xls");
        }

        /** Reference to Name*/
        [Test]
        public void Test13224()
        {
            OpenSample("13224.xls");
        }

        /** Illegal argument exception - cannot store duplicate value in Map*/
        [Test]
        public void Test19599()
        {
            OpenSample("19599-1.xls");
            OpenSample("19599-2.xls");
        }
        [Test]
        public void Test24215()
        {
            HSSFWorkbook wb = OpenSample("24215.xls");

            for (int sheetIndex = 0; sheetIndex < wb.NumberOfSheets; sheetIndex++)
            {
                ISheet sheet = wb.GetSheetAt(sheetIndex);
                int rows = sheet.LastRowNum;

                for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    int cells = row.LastCellNum;

                    for (int cellIndex = 0; cellIndex < cells; cellIndex++)
                    {
                        row.GetCell(cellIndex);
                    }
                }
            }
        }
        [Test]
        [Ignore("not found in poi 3.8beat4")]
        public void Test18800()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            book.CreateSheet("TEST");
            ISheet sheet = book.CloneSheet(0);
            book.SetSheetName(1, "CLONE");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Test"));

            book = WriteOutAndReadBack(book);
            sheet = book.GetSheet("CLONE");
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            Assert.AreEqual("Test", cell.RichStringCellValue.String);
        }

        /**
         * Merged regions were being Removed from the parent in cloned sheets
         */
        [Test]
        [Ignore("this test was not found in poi 3.8beta4")]
        public void Test22720()
        {
            HSSFWorkbook workBook = new HSSFWorkbook();
            workBook.CreateSheet("TEST");
            ISheet template = workBook.GetSheetAt(0);

            template.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2));
            template.AddMergedRegion(new CellRangeAddress(1, 2, 0, 2));

            ISheet clone = workBook.CloneSheet(0);
            int originalMerged = template.NumMergedRegions;
            Assert.AreEqual(2, originalMerged, "2 merged regions");

            //        Remove merged regions from clone
            for (int i = template.NumMergedRegions - 1; i >= 0; i--)
            {
                clone.RemoveMergedRegion(i);
            }

            Assert.AreEqual(originalMerged, template.NumMergedRegions, "Original Sheet's Merged Regions were Removed");
            //        Check if template's merged regions are OK
            if (template.NumMergedRegions > 0)
            {
                // fetch the first merged region...EXCEPTION OCCURS HERE
                template.GetMergedRegion(0);
            }
            //make sure we dont exception

        }

        /**Tests read and Write of Unicode strings in formula results
         * bug and Testcase submitted by Sompop Kumnoonsate
         * The file contains THAI unicode characters.
         */
        [Test]
        public void TestUnicodeStringFormulaRead()
        {

            HSSFWorkbook w = OpenSample("25695.xls");

            ICell a1 = w.GetSheetAt(0).GetRow(0).GetCell(0);
            ICell a2 = w.GetSheetAt(0).GetRow(0).GetCell(1);
            ICell b1 = w.GetSheetAt(0).GetRow(1).GetCell(0);
            ICell b2 = w.GetSheetAt(0).GetRow(1).GetCell(1);
            ICell c1 = w.GetSheetAt(0).GetRow(2).GetCell(0);
            ICell c2 = w.GetSheetAt(0).GetRow(2).GetCell(1);
            ICell d1 = w.GetSheetAt(0).GetRow(3).GetCell(0);
            ICell d2 = w.GetSheetAt(0).GetRow(3).GetCell(1);

            if (false)
            {
#if !HIDE_UNREACHABLE_CODE
                // THAI code page
                Console.WriteLine("a1=" + unicodeString(a1));
                Console.WriteLine("a2=" + unicodeString(a2));
                // US code page
                Console.WriteLine("b1=" + unicodeString(b1));
                Console.WriteLine("b2=" + unicodeString(b2));
                // THAI+US
                Console.WriteLine("c1=" + unicodeString(c1));
                Console.WriteLine("c2=" + unicodeString(c2));
                // US+THAI
                Console.WriteLine("d1=" + unicodeString(d1));
                Console.WriteLine("d2=" + unicodeString(d2));
#endif
            }
            ConfirmSameCellText(a1, a2);
            ConfirmSameCellText(b1, b2);
            ConfirmSameCellText(c1, c2);
            ConfirmSameCellText(d1, d2);

            HSSFWorkbook rw = WriteOutAndReadBack(w);

            ICell ra1 = rw.GetSheetAt(0).GetRow(0).GetCell(0);
            ICell ra2 = rw.GetSheetAt(0).GetRow(0).GetCell(1);
            ICell rb1 = rw.GetSheetAt(0).GetRow(1).GetCell(0);
            ICell rb2 = rw.GetSheetAt(0).GetRow(1).GetCell(1);
            ICell rc1 = rw.GetSheetAt(0).GetRow(2).GetCell(0);
            ICell rc2 = rw.GetSheetAt(0).GetRow(2).GetCell(1);
            ICell rd1 = rw.GetSheetAt(0).GetRow(3).GetCell(0);
            ICell rd2 = rw.GetSheetAt(0).GetRow(3).GetCell(1);

            ConfirmSameCellText(a1, ra1);
            ConfirmSameCellText(b1, rb1);
            ConfirmSameCellText(c1, rc1);
            ConfirmSameCellText(d1, rd1);

            ConfirmSameCellText(a1, ra2);
            ConfirmSameCellText(b1, rb2);
            ConfirmSameCellText(c1, rc2);
            ConfirmSameCellText(d1, rd2);
        }

        private static void ConfirmSameCellText(ICell a, ICell b)
        {
            Assert.AreEqual(a.RichStringCellValue.String, b.RichStringCellValue.String);
        }
        private static String unicodeString(ICell cell)
        {
            String ss = cell.RichStringCellValue.String;
            char[] s = ss.ToCharArray();
            StringBuilder sb = new StringBuilder();
            for (int x = 0; x < s.Length; x++)
            {
                sb.Append("\\u").Append(StringUtil.ToHexString(s[x]));
            }
            return sb.ToString();
        }

        /** Error in Opening wb*/
        [Test]
        public void Test32822()
        {
            OpenSample("32822.xls");
        }
        /**Assert.Fail to read wb with chart */
        [Test]
        public void Test15573()
        {
            OpenSample("15573.xls");
        }

        /**names and macros */
        [Test]
        public void Test27852()
        {
            HSSFWorkbook wb = OpenSample("27852.xls");

            for (int i = 0; i < wb.NumberOfNames; i++)
            {
                NPOI.SS.UserModel.IName name = wb.GetNameAt(i);
                //name.NameName();
                if (name.IsFunctionName)
                {
                    continue;
                }
                //name.Reference;
            }
        }
        [Test]
        [Ignore("not found in poi")]
        public void Test28031()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            String formulaText =
                "IF(ROUND(A2*B2*C2,2)>ROUND(B2*D2,2),ROUND(A2*B2*C2,2),ROUND(B2*D2,2))";
            cell.CellFormula = (formulaText);

            Assert.AreEqual(formulaText, cell.CellFormula);
            WriteTestOutputFileForViewing(wb, "output28031.xls");
        }
        [Test]
        public void Test33082()
        {
            OpenSample("33082.xls");
        }
        [Test]
        public void Test34775()
        {
            try
            {
                OpenSample("34775.xls");
            }
            catch (NullReferenceException)
            {
                throw new AssertionException("identified bug 34775");
            }
        }

        /** Error when reading then writing ArrayValues in NameRecord's*/
        [Test]
        public void Test37630()
        {
            HSSFWorkbook wb = OpenSample("37630.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 25183: org.apache.poi.hssf.usermodel.Sheet.SetPropertiesFromSheet
         */
        [Test]
        public void Test25183()
        {
            HSSFWorkbook wb = OpenSample("25183.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 26100: 128-character message in IF statement cell causes HSSFWorkbook Open Assert.Failure
         */
        [Test]
        public void Test26100()
        {
            HSSFWorkbook wb = OpenSample("26100.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 27933: Unable to use a template (xls) file containing a wmf graphic
         */
        [Test]
        public void Test27933()
        {
            HSSFWorkbook wb = OpenSample("27933.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 29206:      NPE on Sheet.GetRow for blank rows
         */
        [Test]
        public void Test29206()
        {
            //the first Check with blank workbook
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            for (int i = 1; i < 400; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    row.GetCell(0);
                }
            }

            //now Check on an existing xls file
            wb = OpenSample("Simple.xls");

            for (int i = 1; i < 400; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    row.GetCell(0);
                }
            }
        }

        /**
         * Bug 29675: POI 2.5 corrupts output when starting workbook has a graphic
         */
        [Test]
        public void Test29675()
        {
            HSSFWorkbook wb = OpenSample("29675.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 29942: Importing Excel files that have been Created by Open Office on Linux
         */
        [Test]
        public void Test29942()
        {
            HSSFWorkbook wb = OpenSample("29942.xls");

            ISheet sheet = wb.GetSheetAt(0);
            int count = 0;
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell = row.GetCell(0);
                    Assert.AreEqual(CellType.String, cell.CellType);
                    count++;
                }
            }
            Assert.AreEqual(85, count); //should read 85 rows

            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 29982: Unable to read spreadsheet when dropdown list cell is selected -
         *  Unable to construct record instance
         */
        [Test]
        public void Test29982()
        {
            HSSFWorkbook wb = OpenSample("29982.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 30540: Sheet.SetRowBreak throws NullPointerException
         */
        [Test]
        public void Test30540()
        {
            HSSFWorkbook wb = OpenSample("30540.xls");

            ISheet s = wb.GetSheetAt(0);
            s.SetRowBreak(1);
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 31749: {Need help urgently}[This is critical] workbook.Write() corrupts the file......?
         */
        [Test]
        public void Test31749()
        {
            HSSFWorkbook wb = OpenSample("31749.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 31979: {urgent help needed .....}poi library does not support form objects properly.
         */
        public void Test31979()
        {
            HSSFWorkbook wb = OpenSample("31979.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 35564: Cell.java: NullPtrExc in isGridsPrinted() and getProtect()
         *  when HSSFWorkbook is Created from file
         */
        [Test]
        public void Test35564()
        {
            HSSFWorkbook wb = OpenSample("35564.xls");

            ISheet sheet = wb.GetSheetAt(0);
            Assert.AreEqual(false, sheet.IsPrintGridlines);
            Assert.AreEqual(false, sheet.Protect);

            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 35565: Cell.java: NullPtrExc in getColumnBreaks() when HSSFWorkbook is Created from file
         */
        [Test]
        public void Test35565()
        {
            HSSFWorkbook wb = OpenSample("35565.xls");

            ISheet sheet = wb.GetSheetAt(0);
            Assert.IsNotNull(sheet);
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 37376: Cannot Open the saved Excel file if Checkbox controls exceed certain limit
         */
        [Test]
        public void Test37376()
        {
            HSSFWorkbook wb = OpenSample("37376.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 40285:      CellIterator Skips First Column
         */
        [Test]
        public void Test40285()
        {
            HSSFWorkbook wb = OpenSample("40285.xls");

            ISheet sheet = wb.GetSheetAt(0);
            int rownum = 0;
            for (IEnumerator it = sheet.GetRowEnumerator(); it.MoveNext(); rownum++)
            {
                IRow row = (IRow)it.Current;
                Assert.AreEqual(rownum, row.RowNum);
                int cellNum = 0;
                for (IEnumerator it2 = row.GetEnumerator(); it2.MoveNext(); cellNum++)
                {
                    ICell cell = (ICell)it2.Current;
                    Assert.AreEqual(cellNum, cell.ColumnIndex);
                }
            }
        }

        /**
         * Bug 40296:      Cell.SetCellFormula throws
         *   ClassCastException if cell is Created using Row.CreateCell(short column, int type)
         */
        [Test]
        public void Test40296()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            HSSFWorkbook workBook = new HSSFWorkbook();
            ISheet workSheet = workBook.CreateSheet("Sheet1");
            ICell cell;
            IRow row = workSheet.CreateRow(0);
            cell = row.CreateCell(0, CellType.Numeric);
            cell.SetCellValue(1.0);
            cell = row.CreateCell(1, CellType.Numeric);
            cell.SetCellValue(2.0);
            cell = row.CreateCell(2, CellType.Formula);
            cell.CellFormula = ("SUM(A1:B1)");

            WriteOutAndReadBack(wb);
        }

        /**
         * Test bug 38266: NPE when Adding a row break
         *
         * User's diagnosis:
         * 1. Manually (i.e., not using POI) Create an Excel Workbook, making sure it
         * contains a sheet that doesn't have any row breaks
         * 2. Using POI, Create a new HSSFWorkbook from the template in step #1
         * 3. Try Adding a row break (via sheet.SetRowBreak()) to the sheet mentioned in step #1
         * 4. Get a NullPointerException
         */
        [Test]
        public void Test38266()
        {
            String[] files = { "Simple.xls", "SimpleMultiCell.xls", "duprich1.xls" };
            for (int i = 0; i < files.Length; i++)
            {
                HSSFWorkbook wb = OpenSample(files[i]);

                ISheet sheet = wb.GetSheetAt(0);
                int[] breaks = sheet.RowBreaks;
                Assert.AreEqual(0, breaks.Length);

                //Add 3 row breaks
                for (int j = 1; j <= 3; j++)
                {
                    sheet.SetRowBreak(j * 20);
                }
            }
        }
        [Test]
        public void Test40738()
        {
            HSSFWorkbook wb = OpenSample("SimpleWithAutofilter.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 44200: Sheet not cloneable when Note Added to excel cell
         */
        [Test]
        public void Test44200()
        {
            HSSFWorkbook wb = OpenSample("44200.xls");

            wb.CloneSheet(0);
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 44201: Sheet not cloneable when validation Added to excel cell
         */
        [Test]
        public void Test44201()
        {
            HSSFWorkbook wb = OpenSample("44201.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 37684  : Unhandled Continue Record Error
         */
        [Test]
        public void Test37684()
        {
            HSSFWorkbook wb = OpenSample("37684-1.xls");
            WriteOutAndReadBack(wb);


            wb = OpenSample("37684-2.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 41139: Constructing HSSFWorkbook is Assert.Failed,threw threw ArrayIndexOutOfBoundsException for creating UnknownRecord
         */
        [Test]
        public void Test41139()
        {
            HSSFWorkbook wb = OpenSample("41139.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 41546: Constructing HSSFWorkbook is Assert.Failed,
         *  Unknown Ptg in Formula: 0x1a (26)
         */
        [Test]
        public void Test41546()
        {
            HSSFWorkbook wb = OpenSample("41546.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
            wb = WriteOutAndReadBack(wb);
            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * Bug 42564: Some files from Access were giving a RecordFormatException
         *  when reading the BOFRecord
         */
        [Test]
        public void Test42564()
        {
            HSSFWorkbook wb = OpenSample("ex42564-21435.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 42564: Some files from Access also have issues
         *  with the NameRecord, once you get past the BOFRecord
         *  issue.
         */
        [Test]
        public void Test42564Alt()
        {
            HSSFWorkbook wb = OpenSample("ex42564-21503.xls");
            WriteOutAndReadBack(wb);
        }

        /**
         * Bug 42618: RecordFormatException reading a file containing
         *     =CHOOSE(2,A2,A3,A4)
         */
        [Test]
        public void Test42618()
        {
            HSSFWorkbook wb = OpenSample("SimpleWithChoose.xls");
            wb = WriteOutAndReadBack(wb);
            // Check we detect the string properly too
            ISheet s = wb.GetSheetAt(0);

            // Textual value
            IRow r1 = s.GetRow(0);
            ICell c1 = r1.GetCell(1);
            Assert.AreEqual("=CHOOSE(2,A2,A3,A4)", c1.RichStringCellValue.ToString());

            // Formula Value
            IRow r2 = s.GetRow(1);
            ICell c2 = r2.GetCell(1);
            Assert.AreEqual(25, (int)c2.NumericCellValue);

            try
            {
                Assert.AreEqual("CHOOSE(2,A2,A3,A4)", c2.CellFormula);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.StartsWith("Too few arguments")
                        && e.Message.IndexOf("ConcatPtg") > 0)
                {
                    throw new AssertionException("identified bug 44306");
                }
            }
        }

        /**
         * Something up with the FileSharingRecord
         */
        [Test]
        public void Test43251()
        {

            // Used to blow up with an ArgumentException
            //  when creating a FileSharingRecord
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("43251.xls");
            }
            catch (ArgumentException)
            {
                throw new AssertionException("identified bug 43251");
            }

            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * Crystal reports generates files with short
         *  StyleRecords, which is against the spec
         */
        [Test]
        public void Test44471()
        {

            // Used to blow up with an ArrayIndexOutOfBounds
            //  when creating a StyleRecord
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("OddStyleRecord.xls");
            }
            catch (IndexOutOfRangeException)
            {
                throw new AssertionException("Identified bug 44471");
            }

            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * Files with "read only recommended" were giving
         *  grief on the FileSharingRecord
         */
        [Test]
        public void Test44536()
        {

            // Used to blow up with an ArgumentException
            //  when creating a FileSharingRecord
            HSSFWorkbook wb = OpenSample("ReadOnlyRecommended.xls");

            // Check read only advised
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.IsTrue(wb.IsWriteProtected);

            // But also Check that another wb isn't
            wb = OpenSample("SimpleWithChoose.xls");
            Assert.IsFalse(wb.IsWriteProtected);
        }

        /**
         * Some files were having problems with the DVRecord,
         *  probably due to dropdowns
         */
        [Test]
        public void Test44593()
        {

            // Used to blow up with an ArgumentException
            //  when creating a DVRecord
            // Now won't, but no idea if this means we have
            //  rubbish in the DVRecord or not...
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("44593.xls");
            }
            catch (ArgumentException)
            {
                throw new AssertionException("Identified bug 44593");
            }

            Assert.AreEqual(2, wb.NumberOfSheets);
        }

        /**
         * Used to give problems due to trying to read a zero
         *  Length string, but that's now properly handled
         */
        [Test]
        public void Test44643()
        {

            // Used to blow up with an ArgumentException
            HSSFWorkbook wb;
            try
            {
                wb = OpenSample("44643.xls");
            }
            catch (ArgumentException)
            {
                throw new AssertionException("identified bug 44643");
            }

            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
         * User reported the wrong number of rows from the
         *  iterator, but we can't replicate that
         */
        [Test]
        public void Test44693()
        {

            HSSFWorkbook wb = OpenSample("44693.xls");
            ISheet s = wb.GetSheetAt(0);

            // Rows are 1 to 713
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(712, s.LastRowNum);
            Assert.AreEqual(713, s.PhysicalNumberOfRows);

            // Now Check the iterator
            int rowsSeen = 0;
            for (IEnumerator i = s.GetRowEnumerator(); i.MoveNext();)
            {
                IRow r = (IRow)i.Current;
                Assert.IsNotNull(r);
                rowsSeen++;
            }
            Assert.AreEqual(713, rowsSeen);
        }

        /**
         * Bug 28774: Excel will crash when Opening xls-files with images.
         */
        [Test]
        public void Test28774()
        {
            HSSFWorkbook wb = OpenSample("28774.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        /**
         * Had a problem apparently, not sure what as it
         *  works just fine...
         */
        [Test]
        public void Test44891()
        {
            HSSFWorkbook wb = OpenSample("44891.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        /**
         * Bug 44235: Ms Excel can't Open save as excel file
         *
         * Works fine with poi-3.1-beta1.
         */
        [Test]
        public void Test44235()
        {
            HSSFWorkbook wb = OpenSample("44235.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }


        [Test]
        public void Test36947()
        {
            HSSFWorkbook wb = OpenSample("36947.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        [Test]
        public void Test39634()
        {
            HSSFWorkbook wb = OpenSample("39634.xls");
            Assert.IsTrue(true, "no errors reading sample xls");
            WriteOutAndReadBack(wb);
            Assert.IsTrue(true, "no errors writing sample xls");
        }

        /**
         * Problems with extracting Check boxes from
         *  HSSFObjectData
         * @
         */
        [Test]
        public void Test44840()
        {
            HSSFWorkbook wb = OpenSample("WithCheckBoxes.xls");

            // Take a look at the embedded objects
            IList<HSSFObjectData> objects = wb.GetAllEmbeddedObjects();
            Assert.AreEqual(1, objects.Count);

            HSSFObjectData obj = (HSSFObjectData)objects[0];
            Assert.IsNotNull(obj);

            // Peek inside the underlying record
            EmbeddedObjectRefSubRecord rec = obj.FindObjectRecord();
            Assert.IsNotNull(rec);

            //        Assert.AreEqual(32, rec.field_1_stream_id_offset);
            Assert.AreEqual(0, rec.StreamId); // WRONG!
            Assert.AreEqual("Forms.CheckBox.1", rec.OLEClassName);
            Assert.AreEqual(12, rec.ObjectData.Length);

            // Doesn't have a directory
            Assert.IsFalse(obj.HasDirectoryEntry());
            Assert.IsNotNull(obj.GetObjectData());
            Assert.AreEqual(12, obj.GetObjectData().Length);
            Assert.AreEqual("Forms.CheckBox.1", obj.OLE2ClassName);

            try
            {
                obj.GetDirectory();
                Assert.Fail();
            }
            catch (FileNotFoundException)
            {
                // expected during successful Test
            }
        }

        /**
         * Test that we can delete sheets without
         *  breaking the build in named ranges
         *  used for printing stuff.
         */
        [Test]
        public void Test30978()
        {
            HSSFWorkbook wb = OpenSample("30978-alt.xls");
            Assert.AreEqual(1, wb.NumberOfNames);
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Check all names fit within range, and use
            //  DeletedArea3DPtg
            InternalWorkbook w = wb.Workbook;
            for (int i = 0; i < w.NumNames; i++)
            {
                NameRecord r = w.GetNameRecord(i);
                Assert.IsTrue(r.SheetNumber <= wb.NumberOfSheets);

                Ptg[] nd = r.NameDefinition;
                Assert.AreEqual(1, nd.Length);
                Assert.IsTrue(nd[0] is DeletedArea3DPtg);
            }


            // Delete the 2nd sheet
            wb.RemoveSheetAt(1);


            // Re-Check
            Assert.AreEqual(1, wb.NumberOfNames);
            Assert.AreEqual(2, wb.NumberOfSheets);

            for (int i = 0; i < w.NumNames; i++)
            {
                NameRecord r = w.GetNameRecord(i);
                Assert.IsTrue(r.SheetNumber <= wb.NumberOfSheets);

                Ptg[] nd = r.NameDefinition;
                Assert.AreEqual(1, nd.Length);
                Assert.IsTrue(nd[0] is DeletedArea3DPtg);
            }


            // Save and re-load
            wb = WriteOutAndReadBack(wb);
            w = wb.Workbook;

            Assert.AreEqual(1, wb.NumberOfNames);
            Assert.AreEqual(2, wb.NumberOfSheets);

            for (int i = 0; i < w.NumNames; i++)
            {
                NameRecord r = w.GetNameRecord(i);
                Assert.IsTrue(r.SheetNumber <= wb.NumberOfSheets);

                Ptg[] nd = r.NameDefinition;
                Assert.AreEqual(1, nd.Length);
                Assert.IsTrue(nd[0] is DeletedArea3DPtg);
            }
        }

        /**
         * Test that fonts get added properly
         */
        [Test]
        public void Test45338()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Assert.AreEqual(4, wb.NumberOfFonts);

            ISheet s = wb.CreateSheet();
            s.CreateRow(0);
            s.CreateRow(1);
            ICell c1 = s.GetRow(0).CreateCell(0);
            ICell c2 = s.GetRow(1).CreateCell(0);

            Assert.AreEqual(4, wb.NumberOfFonts);

            IFont f1 = wb.GetFontAt((short)0);
            Assert.AreEqual(400, f1.Boldweight);

            // Check that asking for the same font
            //  multiple times gives you the same thing.
            // Otherwise, our Tests wouldn't work!
            Assert.AreEqual(
                    wb.GetFontAt((short)0),
                    wb.GetFontAt((short)0)
            );
            Assert.AreEqual(
                    wb.GetFontAt((short)2),
                    wb.GetFontAt((short)2)
            );
            Assert.IsTrue(
                    wb.GetFontAt((short)0)
                    !=
                    wb.GetFontAt((short)2)
            );

            // Look for a new font we have
            //  yet to Add
            Assert.IsNull(
                wb.FindFont(
                    (short)11, (short)123, (short)22,
                    "Thingy", false, true, FontSuperScript.Sub, FontUnderlineType.Double
                )
            );

            IFont nf = wb.CreateFont();
            Assert.AreEqual(5, wb.NumberOfFonts);

            Assert.AreEqual(5, nf.Index);
            Assert.AreEqual(nf, wb.GetFontAt((short)5));

            nf.Boldweight = ((short)11);
            nf.Color = ((short)123);
            nf.FontHeight = ((short)22);
            nf.FontName = ("Thingy");
            nf.IsItalic = (false);
            nf.IsStrikeout = (true);
            nf.TypeOffset = FontSuperScript.Sub;
            nf.Underline = FontUnderlineType.Double;

            Assert.AreEqual(5, wb.NumberOfFonts);
            Assert.AreEqual(nf, wb.GetFontAt((short)5));

            // Find it now
            Assert.IsNotNull(
                wb.FindFont(
                    (short)11, (short)123, (short)22,
                    "Thingy", false, true, FontSuperScript.Sub, FontUnderlineType.Double
                )
            );
            Assert.AreEqual(
                5,
                wb.FindFont(
                       (short)11, (short)123, (short)22,
                       "Thingy", false, true, FontSuperScript.Sub, FontUnderlineType.Double
                   ).Index
            );
            Assert.AreEqual(nf,
                   wb.FindFont(
                       (short)11, (short)123, (short)22,
                       "Thingy", false, true, FontSuperScript.Sub, FontUnderlineType.Double
                   )
            );

            wb.Close();
        }

        /**
         * From the mailing list - ensure we can handle a formula
         *  containing a zip code, eg ="70164"
         */
        [Test]
        public void TestZipCodeFormulas()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();
            s.CreateRow(0);
            ICell c1 = s.GetRow(0).CreateCell(0);
            ICell c2 = s.GetRow(0).CreateCell(1);
            ICell c3 = s.GetRow(0).CreateCell(2);

            // As number and string
            c1.CellFormula = ("70164");
            c2.CellFormula = ("\"70164\"");
            c3.CellFormula = ("\"90210\"");

            // Check the formulas
            Assert.AreEqual("70164", c1.CellFormula);
            Assert.AreEqual("\"70164\"", c2.CellFormula);

            // And Check the values - blank
            ConfirmCachedValue(0.0, c1);
            ConfirmCachedValue(0.0, c2);
            ConfirmCachedValue(0.0, c3);

            // Try changing the cached value on one of the string
            //  formula cells, so we can see it updates properly
            c3.SetCellValue(new HSSFRichTextString("Test"));
            ConfirmCachedValue("Test", c3);
            try
            {
                double a = c3.NumericCellValue;
                throw new AssertionException("exception should have been thrown");
            }
            catch (InvalidOperationException e)
            {
                Assert.AreEqual("Cannot get a numeric value from a text formula cell", e.Message);
            }


            // Now Evaluate, they should all be changed
            HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);
            eval.EvaluateFormulaCell(c1);
            eval.EvaluateFormulaCell(c2);
            eval.EvaluateFormulaCell(c3);

            // Check that the cells now contain
            //  the correct values
            ConfirmCachedValue(70164.0, c1);
            ConfirmCachedValue("70164", c2);
            ConfirmCachedValue("90210", c3);


            // Write and read
            HSSFWorkbook nwb = WriteOutAndReadBack(wb);
            HSSFSheet ns = (HSSFSheet)nwb.GetSheetAt(0);
            ICell nc1 = ns.GetRow(0).GetCell(0);
            ICell nc2 = ns.GetRow(0).GetCell(1);
            ICell nc3 = ns.GetRow(0).GetCell(2);

            // Re-Check
            ConfirmCachedValue(70164.0, nc1);
            ConfirmCachedValue("70164", nc2);
            ConfirmCachedValue("90210", nc3);

            int i = 0;
            for (IEnumerator<CellValueRecordInterface> it = ns.Sheet.GetCellValueIterator(); it.MoveNext(); i++)
            {
                CellValueRecordInterface cvr = it.Current;
                if (cvr is FormulaRecordAggregate)
                {
                    FormulaRecordAggregate fr = (FormulaRecordAggregate)cvr;

                    if (i == 0)
                    {
                        Assert.AreEqual(70164.0, fr.FormulaRecord.Value, 0.0001);
                        Assert.IsNull(fr.StringRecord);
                    }
                    else if (i == 1)
                    {
                        Assert.AreEqual(0.0, fr.FormulaRecord.Value, 0.0001);
                        Assert.IsNotNull(fr.StringRecord);
                        Assert.AreEqual("70164", fr.StringRecord.String);
                    }
                    else
                    {
                        Assert.AreEqual(0.0, fr.FormulaRecord.Value, 0.0001);
                        Assert.IsNotNull(fr.StringRecord);
                        Assert.AreEqual("90210", fr.StringRecord.String);
                    }
                }
            }
            Assert.AreEqual(3, i);
        }

        private static void ConfirmCachedValue(double expectedValue, ICell cell)
        {
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(CellType.Numeric, cell.CachedFormulaResultType);
            Assert.AreEqual(expectedValue, cell.NumericCellValue, 0.0);
        }
        private static void ConfirmCachedValue(String expectedValue, ICell cell)
        {
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(CellType.String, cell.CachedFormulaResultType);
            Assert.AreEqual(expectedValue, cell.RichStringCellValue.String);
        }

        /**
         * Problem with "Vector Rows", eg a whole
         *  column which is set to the result of
         *  {=sin(B1:B9)}(9,1), so that each cell is
         *  shown to have the contents
         *  {=sin(B1:B9){9,1)[rownum][0]
         * In this sample file, the vector column
         *  is C, and the data column is B.
         *
         * For now, blows up with an exception from ExtPtg
         *  Expected ExpPtg to be converted from Shared to Non-Shared...
         */
        [Test]
        [Ignore("this test is disabled in poi.")]
        public void DISABLED_Test43623()
        {
            HSSFWorkbook wb = OpenSample("43623.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);

            ISheet s1 = wb.GetSheetAt(0);

            ICell c1 = s1.GetRow(0).GetCell(2);
            ICell c2 = s1.GetRow(1).GetCell(2);
            ICell c3 = s1.GetRow(2).GetCell(2);

            // These formula contents are a guess...
            Assert.AreEqual("{=sin(B1:B9){9,1)[0][0]", c1.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[1][0]", c2.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[2][0]", c3.CellFormula);

            // Save and re-Open, ensure it still works
            HSSFWorkbook nwb = WriteOutAndReadBack(wb);
            ISheet ns1 = nwb.GetSheetAt(0);
            ICell nc1 = ns1.GetRow(0).GetCell(2);
            ICell nc2 = ns1.GetRow(1).GetCell(2);
            ICell nc3 = ns1.GetRow(2).GetCell(2);

            Assert.AreEqual("{=sin(B1:B9){9,1)[0][0]", nc1.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[1][0]", nc2.CellFormula);
            Assert.AreEqual("{=sin(B1:B9){9,1)[2][0]", nc3.CellFormula);
        }

        /**
         * People are all getting confused about the last
         *  row and cell number
         */
        [Test]
        public void Test30635()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();

            // No rows, everything is 0
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(0, s.LastRowNum);
            Assert.AreEqual(0, s.PhysicalNumberOfRows);

            // One row, most things are 0, physical is 1
            s.CreateRow(0);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(0, s.LastRowNum);
            Assert.AreEqual(1, s.PhysicalNumberOfRows);

            // And another, things change
            s.CreateRow(4);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(4, s.LastRowNum);
            Assert.AreEqual(2, s.PhysicalNumberOfRows);


            // Now start on cells
            IRow r = s.GetRow(0);
            Assert.AreEqual(-1, r.FirstCellNum);
            Assert.AreEqual(-1, r.LastCellNum);
            Assert.AreEqual(0, r.PhysicalNumberOfCells);

            // Add a cell, things move off -1
            r.CreateCell(0);
            Assert.AreEqual(0, r.FirstCellNum);
            Assert.AreEqual(1, r.LastCellNum); // last cell # + 1
            Assert.AreEqual(1, r.PhysicalNumberOfCells);

            r.CreateCell(1);
            Assert.AreEqual(0, r.FirstCellNum);
            Assert.AreEqual(2, r.LastCellNum); // last cell # + 1
            Assert.AreEqual(2, r.PhysicalNumberOfCells);

            r.CreateCell(4);
            Assert.AreEqual(0, r.FirstCellNum);
            Assert.AreEqual(5, r.LastCellNum); // last cell # + 1
            Assert.AreEqual(3, r.PhysicalNumberOfCells);

            wb.Close();
        }

        /**
         * Data Tables - ptg 0x2
         */
        [Test]
        public void Test44958()
        {
            HSSFWorkbook wb = OpenSample("44958.xls");
            ISheet s;
            IRow r;
            ICell c;

            // Check the contents of the formulas

            // E4 to G9 of sheet 4 make up the table
            s = wb.GetSheet("OneVariable Table Completed");
            r = s.GetRow(3);
            c = r.GetCell(4);
            Assert.AreEqual(CellType.Formula, c.CellType);

            // TODO - Check the formula once tables and
            //  arrays are properly supported


            // E4 to H9 of sheet 5 make up the table
            s = wb.GetSheet("TwoVariable Table Example");
            r = s.GetRow(3);
            c = r.GetCell(4);
            Assert.AreEqual(CellType.Formula, c.CellType);

            // TODO - Check the formula once tables and
            //  arrays are properly supported
        }

        /**
         * 45322: Sheet.autoSizeColumn Assert.Fails when style.GetDataFormat() returns -1
         */
        [Test]
        public void Test45322()
        {
            HSSFWorkbook wb = OpenSample("44958.xls");
            ISheet sh = wb.GetSheetAt(0);
            for (short i = 0; i < 30; i++) sh.AutoSizeColumn(i);
        }

        /**
         * We used to Add too many UncalcRecords to sheets
         *  with diagrams on. Don't any more
         */
        [Test]
        public void Test45414()
        {
            HSSFWorkbook wb = OpenSample("WithThreeCharts.xls");
            wb.GetSheetAt(0).ForceFormulaRecalculation = (true);
            wb.GetSheetAt(1).ForceFormulaRecalculation = (false);
            wb.GetSheetAt(2).ForceFormulaRecalculation = (true);

            // Write out and back in again
            // This used to break
            HSSFWorkbook nwb = WriteOutAndReadBack(wb);

            // Check now set as it should be
            Assert.IsTrue(nwb.GetSheetAt(0).ForceFormulaRecalculation);
            Assert.IsFalse(nwb.GetSheetAt(1).ForceFormulaRecalculation);
            Assert.IsTrue(nwb.GetSheetAt(2).ForceFormulaRecalculation);
        }

        /**
         * Very hidden sheets not displaying as such
         */
        [Test]
        public void Test45761()
        {
            HSSFWorkbook wb = OpenSample("45761.xls");
            Assert.AreEqual(3, wb.NumberOfSheets);

            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));
            Assert.IsTrue(wb.IsSheetHidden(1));
            Assert.IsFalse(wb.IsSheetVeryHidden(1));
            Assert.IsFalse(wb.IsSheetHidden(2));
            Assert.IsTrue(wb.IsSheetVeryHidden(2));

            // Change 0 to be very hidden, and re-load
            wb.SetSheetHidden(0, 2);

            HSSFWorkbook nwb = WriteOutAndReadBack(wb);

            Assert.IsFalse(nwb.IsSheetHidden(0));
            Assert.IsTrue(nwb.IsSheetVeryHidden(0));
            Assert.IsTrue(nwb.IsSheetHidden(1));
            Assert.IsFalse(nwb.IsSheetVeryHidden(1));
            Assert.IsFalse(nwb.IsSheetHidden(2));
            Assert.IsTrue(nwb.IsSheetVeryHidden(2));
        }

        ///**
        // * header / footer text too long
        // */
        //[Test]
        //public void Test45777()
        //{
        //    HSSFWorkbook wb = new HSSFWorkbook();
        //    Sheet s = wb.CreateSheet();

        //    String s248 = "";
        //    for (int i = 0; i < 248; i++)
        //    {
        //        s248 += "x";
        //    }
        //    String s249 = s248 + "1";
        //    String s250 = s248 + "12";
        //    String s251 = s248 + "123";
        //    Assert.AreEqual(248, s248.Length);
        //    Assert.AreEqual(249, s249.Length);
        //    Assert.AreEqual(250, s250.Length);
        //    Assert.AreEqual(251, s251.Length);


        //    // Try on headers
        //    s.Header.Center = (s248);
        //    Assert.AreEqual(254, s.Header.RawHeader.Length);
        //    WriteOutAndReadBack(wb);

        //    s.Header.Center = (s249);
        //    Assert.AreEqual(255, s.Header.RawHeader.Length);
        //    WriteOutAndReadBack(wb);

        //    try
        //    {
        //        s.Header.Center = (s250); // 256
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }

        //    try
        //    {
        //        s.Header.Center = (s251); // 257
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }


        //    // Now try on footers
        //    s.Footer.Center = (s248);
        //    Assert.AreEqual(254, s.Footer.RawFooter.Length);
        //    WriteOutAndReadBack(wb);

        //    s.Footer.Center = (s249);
        //    Assert.AreEqual(255, s.Footer.RawFooter.Length);
        //    WriteOutAndReadBack(wb);

        //    try
        //    {
        //        s.Footer.Center = (s250); // 256
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }

        //    try
        //    {
        //        s.Footer.Center = (s251); // 257
        //        Assert.Fail();
        //    }
        //    catch (ArgumentException e) { }
        //}

        /**
         * Charts with long titles
         */
        [Test]
        public void Test45784()
        {
            // This used to break
            HSSFWorkbook wb = OpenSample("45784.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
          * Cell background colours
          */
        [Test]
        public void Test45492()
        {
            HSSFWorkbook wb = OpenSample("45492.xls");
            ISheet s = wb.GetSheetAt(0);
            IRow r = s.GetRow(0);
            HSSFPalette p = wb.GetCustomPalette();

            ICell auto = r.GetCell(0);
            ICell grey = r.GetCell(1);
            ICell red = r.GetCell(2);
            ICell blue = r.GetCell(3);
            ICell green = r.GetCell(4);

            Assert.AreEqual(64, auto.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, auto.CellStyle.FillBackgroundColor);
            Assert.AreEqual("0:0:0", p.GetColor(64).GetHexString());

            Assert.AreEqual(22, grey.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, grey.CellStyle.FillBackgroundColor);
            Assert.AreEqual("C0C0:C0C0:C0C0", p.GetColor(22).GetHexString());

            Assert.AreEqual(10, red.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, red.CellStyle.FillBackgroundColor);
            Assert.AreEqual("FFFF:0:0", p.GetColor(10).GetHexString());

            Assert.AreEqual(12, blue.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, blue.CellStyle.FillBackgroundColor);
            Assert.AreEqual("0:0:FFFF", p.GetColor(12).GetHexString());

            Assert.AreEqual(11, green.CellStyle.FillForegroundColor);
            Assert.AreEqual(64, green.CellStyle.FillBackgroundColor);
            Assert.AreEqual("0:FFFF:0", p.GetColor(11).GetHexString());
        }
        /**
 * ContinueRecord after EOF
 */
        [Test]
        public void Test46137()
        {
            // This used to break
            HSSFWorkbook wb = OpenSample("46137.xls");
            Assert.AreEqual(7, wb.NumberOfSheets);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }
        /**
         * Newly created sheets need to get a 
         *  proper TabID, otherwise print setup
         *  gets confused on them.
         * Also ensure that print setup refs are
         *  by reference not value 
         */
        [Test]
        public void Test46664()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            ISheet sheet = wb1.CreateSheet("new_sheet");
            IRow row = sheet.CreateRow((short)0);
            row.CreateCell(0).SetCellValue(new HSSFRichTextString("Column A"));
            row.CreateCell(1).SetCellValue(new HSSFRichTextString("Column B"));
            row.CreateCell(2).SetCellValue(new HSSFRichTextString("Column C"));
            row.CreateCell(3).SetCellValue(new HSSFRichTextString("Column D"));
            row.CreateCell(4).SetCellValue(new HSSFRichTextString("Column E"));
            row.CreateCell(5).SetCellValue(new HSSFRichTextString("Column F"));

            //set print area from column a to column c (on first row)
            wb1.SetPrintArea(
                    0, //sheet index
                    0, //start column
                    2, //end column
                    0, //start row
                    0  //end row
            );

            HSSFWorkbook wb2 = WriteOutAndReadBack(wb1);
            wb1.Close();

            // Ensure the tab index
            TabIdRecord tr = null;
            foreach (Record r in wb2.Workbook.Records)
            {
                if (r is TabIdRecord)
                {
                    tr = (TabIdRecord)r;
                }
            }
            Assert.IsNotNull(tr);
            Assert.AreEqual(1, tr._tabids.Length);
            Assert.AreEqual(0, tr._tabids[0]);

            // Ensure the print setup
            Assert.AreEqual("new_sheet!$A$1:$C$1", wb2.GetPrintArea(0));
            Assert.AreEqual("new_sheet!$A$1:$C$1", wb2.GetName("Print_Area").RefersToFormula);

            // Needs reference not value
            NameRecord nr = wb2.Workbook.GetNameRecord(
                  wb2.GetNameIndex("Print_Area")
            );
            Assert.AreEqual("Print_Area", nr.NameText);
            Assert.AreEqual(1, nr.NameDefinition.Length);
            Assert.AreEqual(
                  "new_sheet!$A$1:$C$1",
                  ((Area3DPtg)nr.NameDefinition[0]).ToFormulaString(HSSFEvaluationWorkbook.Create(wb2))
            );

            Assert.AreEqual('R', nr.NameDefinition[0].RVAType);

            wb2.Close();
        }
        /**
         * Odd POIFS blocks issue:
         * block[ 44 ] already removed from org.apache.poi.poifs.storage.BlockListImpl.remove
         */
        [Test]
        public void Test45290()
        {
            HSSFWorkbook wb = OpenSample("45290.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
        }
        /**
 * In POI-2.5 user reported exception when parsing a name with a custom VBA function:
 *  =MY_VBA_FUNCTION("lskdjflsk")
 */
        [Test]
        public void Test30070()
        {
            HSSFWorkbook wb = OpenSample("30070.xls"); //contains custom VBA function 'Commission'
            HSSFSheet sh = (HSSFSheet)wb.GetSheetAt(0);
            HSSFCell cell = (HSSFCell)sh.GetRow(0).GetCell(1);

            //B1 uses VBA in the formula
            Assert.AreEqual("Commission(A1)", cell.CellFormula);

            //name sales_1 refers to Commission(Sheet0!$A$1)
            int idx = wb.GetNameIndex("sales_1");
            Assert.IsTrue(idx != -1);

            HSSFName name = (HSSFName)wb.GetNameAt(idx);
            Assert.AreEqual("Commission(Sheet0!$A$1)", name.RefersToFormula);

        }

        /**
 * The link formulas which is referring to other books cannot be taken (the bug existed prior to POI-3.2)
 * Expected:
 *
 * [link_sub.xls]Sheet1!$A$1
 * [link_sub.xls]Sheet1!$A$2
 * [link_sub.xls]Sheet1!$A$3
 *
 * POI-3.1 output:
 *
 * Sheet1!$A$1
 * Sheet1!$A$2
 * Sheet1!$A$3
 *
 */
        [Test]
        public void Test27364()
        {
            HSSFWorkbook wb = OpenSample("27364.xls");
            ISheet sheet = wb.GetSheetAt(0);

            Assert.AreEqual("[link_sub.xls]Sheet1!$A$1", sheet.GetRow(0).GetCell(0).CellFormula);
            Assert.AreEqual("[link_sub.xls]Sheet1!$A$2", sheet.GetRow(1).GetCell(0).CellFormula);
            Assert.AreEqual("[link_sub.xls]Sheet1!$A$3", sheet.GetRow(2).GetCell(0).CellFormula);
        }

        /**
         * Similar to bug#27364:
         * HSSFCell.GetCellFormula() fails with references to external workbooks
         */
        [Test]
        public void Test31661()
        {
            HSSFWorkbook wb = OpenSample("31661.xls");
            ISheet sheet = wb.GetSheetAt(0);
            ICell cell = sheet.GetRow(11).GetCell(10); //K11
            Assert.AreEqual("+'[GM Budget.xls]8085.4450'!$B$2", cell.CellFormula);
        }

        /**
         * Incorrect handling of non-ISO 8859-1 characters in Windows ANSII Code Page 1252
         */
        [Test]
        public void Test27394()
        {
            HSSFWorkbook wb = OpenSample("27394.xls");
            Assert.AreEqual("\u0161\u017E", wb.GetSheetName(0));
            Assert.AreEqual("\u0161\u017E\u010D\u0148\u0159", wb.GetSheetName(1));
            ISheet sheet = wb.GetSheetAt(0);

            Assert.AreEqual("\u0161\u017E", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("\u0161\u017E\u010D\u0148\u0159", sheet.GetRow(1).GetCell(0).StringCellValue);
        }

        /**
         * Multiple calls of HSSFWorkbook.Write result in corrupted xls
         */
        [Test]
        public void Test32191()
        {
            HSSFWorkbook wb = OpenSample("27394.xls");

            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);
            long size1 = out1.Length;
            out1.Close();

            out1 = new MemoryStream();
            wb.Write(out1);
            long size2 = out1.Length;
            out1.Close();

            Assert.AreEqual(size1, size2);
            out1 = new MemoryStream();
            wb.Write(out1);
            long size3 = out1.Length;
            out1.Close();
            Assert.AreEqual(size2, size3);

        }

        /**
         * java.io.IOException: block[ 0 ] already removed
         * (is an excel 95 file though)
         */
        [Test]
        public void Test46904()
        {
            try
            {
                //OpenSample("46904.xls");
                OPOIFSFileSystem fs = new OPOIFSFileSystem(
                    HSSFITestDataProvider.Instance.OpenWorkbookStream("46904.xls"));
                new HSSFWorkbook(fs.Root, false).Close();
                Assert.Fail();
            }
            catch (OldExcelFormatException e)
            {
                Assert.IsTrue(e.Message.StartsWith(
                        "The supplied spreadsheet seems to be Excel"
                ));
            }

            try
            {
                NPOIFSFileSystem fs = new NPOIFSFileSystem(
                        HSSFITestDataProvider.Instance.OpenWorkbookStream("46904.xls"));
                try
                {
                    new HSSFWorkbook(fs.Root, false).Close();
                    Assert.Fail();
                }
                finally
                {
                    fs.Close();
                }
            }
            catch (OldExcelFormatException e)
            {
                Assert.IsTrue(e.Message.StartsWith(
                        "The supplied spreadsheet seems to be Excel"
                ));
            }

        }

        /**
         * java.lang.NegativeArraySizeException reading long
         *  non-unicode data for a name record
         */
        [Test]
        public void Test47034()
        {
            HSSFWorkbook wb = OpenSample("47034.xls");
            Assert.AreEqual(893, wb.NumberOfNames);
            Assert.AreEqual("Matthew\\Matthew11_1\\Matthew2331_1\\Matthew2351_1\\Matthew2361_1___lab", wb.GetNameName(300));
        }

        /**
         * HSSFRichTextString.Length returns negative for really long strings.
         * The Test file was created in OpenOffice 3.0 as Excel does not allow cell text longer than 32,767 characters
         */
        [Test]
        public void Test46368()
        {
            HSSFWorkbook wb = OpenSample("46368.xls");
            ISheet s = wb.GetSheetAt(0);
            ICell cell1 = s.GetRow(0).GetCell(0);
            Assert.AreEqual(32770, cell1.StringCellValue.Length);

            ICell cell2 = s.GetRow(2).GetCell(0);
            Assert.AreEqual(32766, cell2.StringCellValue.Length);
        }

        /**
         * Short records on certain sheets with charts in them
         */
        [Test]
        public void Test48180()
        {
            HSSFWorkbook wb = OpenSample("48180.xls");

            ISheet s = wb.GetSheetAt(0);
            ICell cell1 = s.GetRow(0).GetCell(0);
            Assert.AreEqual("test ", cell1.StringCellValue);

            ICell cell2 = s.GetRow(0).GetCell(1);
            Assert.AreEqual(1.0, cell2.NumericCellValue);
        }

        /**
         * POI 3.5 beta 7 can not read excel file contain list box (Form Control)  
         */
        [Test]
        public void Test47701()
        {
            OpenSample("47701.xls");
        }
        /**
 * Test for a file with NameRecord with NameCommentRecord comments
 */
        public void Test49185()
        {
            HSSFWorkbook wb = OpenSample("49185.xls");
            IName name = wb.GetName("foobarName");
            Assert.AreEqual("This is a comment", name.Comment);

            // Rename the name, comment comes with it
            name.NameName = ("ChangedName");
            Assert.AreEqual("This is a comment", name.Comment);

            // Save and re-check
            wb = WriteOutAndReadBack(wb);
            name = wb.GetName("ChangedName");
            Assert.AreEqual("This is a comment", name.Comment);

            // Now try to change it
            name.Comment = ("Changed Comment");
            Assert.AreEqual("Changed Comment", name.Comment);

            // Save and re-check
            wb = WriteOutAndReadBack(wb);
            name = wb.GetName("ChangedName");
            Assert.AreEqual("Changed Comment", name.Comment);
        }
        /// <summary>
        /// http://npoi.codeplex.com/WorkItem/View.aspx?WorkItemId=5010
        /// </summary>
        [Test]
        public void TestNPOIBug5010()
        {
            try
            {
                OpenSample("NPOIBug5010.xls");
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Contains("Unable to construct record instance"))
                {
                    throw new AssertionException("identified NPOI bug 5010");
                }
            }
        }
        /// <summary>
        /// http://npoi.codeplex.com/WorkItem/View.aspx?WorkItemId=5139
        /// </summary>
        [Test]
        public void TestNPOIBug5139()
        {
            try
            {
                OpenSample("NPOIBug5139.xls");
            }
            catch (LeftoverDataException e)
            {
                if (e.Message.StartsWith("Initialisation of record 0x862"))
                {
                    throw new AssertionException("identified NPOI bug 5139");
                }
            }
        }
        /**
 * Vertically aligned text
 */
        [Test]
        public void Test49524()
        {
            HSSFWorkbook wb = OpenSample("49524.xls");
            ISheet s = wb.GetSheetAt(0);
            IRow r = s.GetRow(0);
            ICell rotated = r.GetCell(0);
            ICell normal = r.GetCell(1);

            // Check the current ones
            Assert.AreEqual(0, normal.CellStyle.Rotation);
            Assert.AreEqual(0xff, rotated.CellStyle.Rotation);

            // Add a new style, also rotated
            ICellStyle cs = wb.CreateCellStyle();
            cs.Rotation = ((short)0xff);
            ICell nc = r.CreateCell(2);
            nc.SetCellValue("New Rotated Text");
            nc.CellStyle = (cs);
            Assert.AreEqual(0xff, nc.CellStyle.Rotation);

            // Write out and read back
            wb = WriteOutAndReadBack(wb);

            // Re-check
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            rotated = r.GetCell(0);
            normal = r.GetCell(1);
            nc = r.GetCell(2);

            Assert.AreEqual(0, normal.CellStyle.Rotation);
            Assert.AreEqual(0xff, rotated.CellStyle.Rotation);
            Assert.AreEqual(0xff, nc.CellStyle.Rotation);
        }

        /**
         * Regression with the PageSettingsBlock
         */
        [Test]
        public void Test49931()
        {
            HSSFWorkbook wb = OpenSample("49931.xls");

            Assert.AreEqual(1, wb.NumberOfSheets);
            Assert.AreEqual("Foo", wb.GetSheetAt(0).GetRow(0).GetCell(0).RichStringCellValue.String);
        }

        /**
         * Missing left/right/centre options on a footer
         */
        [Test]
        public void Test48325()
        {
            HSSFWorkbook wb = OpenSample("48325.xls");
            ISheet sh = wb.GetSheetAt(0);
            IFooter f = sh.Footer;

            // Will show as the centre, as that is what excel does
            //  with an invalid footer lacking left/right/centre details
            Assert.AreEqual("", f.Left, "Left text should be empty");
            Assert.AreEqual("", f.Right, "Right text should be empty");
            Assert.AreEqual(
                  "BlahBlah blah blah  ", f.Center, "Center text should contain the illegal value"
            );
        }

        /**
         * InvalidOperationException received when creating Data validation in sheet with macro
         */
        [Test]
        public void Test50020()
        {
            HSSFWorkbook wb = OpenSample("50020.xls");
            WriteOutAndReadBack(wb);
        }
        [Test]
        public void TestAutoSize_bug50681()
        {
            IWorkbook wb = new HSSFWorkbook();
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
 * Mixture of Ascii and Unicode strings in a 
 *  NameComment record
 */
        [Test]
        public void Test51143()
        {
            HSSFWorkbook wb = OpenSample("51143.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
            wb = WriteOutAndReadBack(wb);
            Assert.AreEqual(1, wb.NumberOfSheets);
        }

        /**
     * The resolution for bug 45777 assumed that the maximum text length in a header / footer
     * record was 256 bytes.  This assumption appears to be wrong.  Since the fix for bug 47244,
     * POI now supports header / footer text lengths beyond 256 bytes.
     */
        [Test]
        public void Test45777()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();

            char[] cc248 = new char[248];
            Arrays.Fill(cc248, 'x');
            String s248 = new String(cc248);

            String s249 = s248 + "1";
            String s250 = s248 + "12";
            String s251 = s248 + "123";
            Assert.AreEqual(248, s248.Length);
            Assert.AreEqual(249, s249.Length);
            Assert.AreEqual(250, s250.Length);
            Assert.AreEqual(251, s251.Length);


            // Try on headers
            s.Header.Center = (s248);
            Assert.AreEqual(254, ((HSSFHeader)s.Header).RawText.Length);
            WriteOutAndReadBack(wb);

            s.Header.Center = (s251);
            Assert.AreEqual(257, ((HSSFHeader)s.Header).RawText.Length);
            WriteOutAndReadBack(wb);

            try
            {
                s.Header.Center = (s250); // 256 bytes required
            }
            catch (ArgumentException)
            {
                throw new AssertionException("Identified bug 47244b - header can be more than 256 bytes");
            }

            try
            {
                s.Header.Center = (s251); // 257 bytes required
            }
            catch (ArgumentException)
            {
                throw new AssertionException("Identified bug 47244b - header can be more than 256 bytes");
            }

            // Now try on footers
            s.Footer.Center = (s248);
            Assert.AreEqual(254, ((HSSFFooter)s.Footer).RawText.Length);
            WriteOutAndReadBack(wb);

            s.Footer.Center = (s251);
            Assert.AreEqual(257, ((HSSFFooter)s.Footer).RawText.Length);
            WriteOutAndReadBack(wb);

            try
            {
                s.Footer.Center = (s250); // 256 bytes required
            }
            catch (ArgumentException)
            {
                throw new AssertionException("Identified bug 47244b - footer can be more than 256 bytes");
            }

            try
            {
                s.Footer.Center = (s251); // 257 bytes required
            }
            catch (ArgumentException)
            {
                throw new AssertionException("Identified bug 47244b - footer can be more than 256 bytes");
            }
        }
        /**
     * Problems with formula references to 
     *  sheets via URLs
     */
        [Test]
        public void Test45970()
        {
            HSSFWorkbook wb = OpenSample("FormulaRefs.xls");
            Assert.AreEqual(3, wb.NumberOfSheets);

            ISheet s = wb.GetSheetAt(0);
            IRow row;

            row = s.GetRow(0);
            Assert.AreEqual(CellType.Numeric, row.GetCell(1).CellType);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            row = s.GetRow(1);
            Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
            Assert.AreEqual("B1", row.GetCell(1).CellFormula);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            row = s.GetRow(2);
            Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
            Assert.AreEqual("Sheet1!B1", row.GetCell(1).CellFormula);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            row = s.GetRow(3);
            Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
            Assert.AreEqual("[Formulas2.xls]Sheet1!B2", row.GetCell(1).CellFormula);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            row = s.GetRow(4);
            Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
            Assert.AreEqual("'[$http://gagravarr.org/FormulaRefs.xls]Sheet1'!B1", row.GetCell(1).CellFormula);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            // Change 4
            row.GetCell(1).CellFormula = ("'[$http://gagravarr.org/FormulaRefs2.xls]Sheet1'!B2");
            row.GetCell(1).SetCellValue(123.0);

            // Add 5
            row = s.CreateRow(5);
            row.CreateCell(1, CellType.Formula);
            row.GetCell(1).CellFormula = ("'[$http://example.com/FormulaRefs.xls]Sheet1'!B1");
            row.GetCell(1).SetCellValue(234.0);


            // Re-test
            wb = WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);

            row = s.GetRow(0);
            Assert.AreEqual(CellType.Numeric, row.GetCell(1).CellType);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            row = s.GetRow(1);
            Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
            Assert.AreEqual("B1", row.GetCell(1).CellFormula);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            row = s.GetRow(2);
            Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
            Assert.AreEqual("Sheet1!B1", row.GetCell(1).CellFormula);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

            row = s.GetRow(3);
            Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
            Assert.AreEqual("[Formulas2.xls]Sheet1!B2", row.GetCell(1).CellFormula);
            Assert.AreEqual(112.0, row.GetCell(1).NumericCellValue);

#if !HIDE_UNREACHABLE_CODE
            // TODO - Fix these so they work...
            if (1 == 2)
            {
                row = s.GetRow(4);
                Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
                Assert.AreEqual("'[\u0005$http://gagravarr.org/FormulaRefs2.xls]Sheet1'!B2", row.GetCell(1).CellFormula);
                Assert.AreEqual(123.0, row.GetCell(1).NumericCellValue);

                row = s.GetRow(5);
                Assert.AreEqual(CellType.Formula, row.GetCell(1).CellType);
                Assert.AreEqual("'[\u0005$http://example.com/FormulaRefs.xls]Sheet1'!B1", row.GetCell(1).CellFormula);
                Assert.AreEqual(234.0, row.GetCell(1).NumericCellValue);
            }
#endif
        }
        [Test]
        public void Test47251()
        {
            // Firstly, try with one that triggers on InterfaceHdrRecord
            OpenSample("47251.xls");

            // Now with one that triggers on NoteRecord
            OpenSample("47251_1.xls");
        }
        /**
     * Round trip a file with an unusual UnicodeString/ExtRst record parts
     */
        [Test]
        public void Test47847()
        {
            HSSFWorkbook wb = OpenSample("47847.xls");
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Find the SST record
            UnicodeString withExt = wb.Workbook.GetSSTString(0);
            UnicodeString withoutExt = wb.Workbook.GetSSTString(31);

            Assert.AreEqual("O:Alloc:Qty", withExt.String);
            Assert.IsTrue((withExt.OptionFlags & 0x0004) == 0x0004);

            Assert.AreEqual("RT", withoutExt.String);
            Assert.IsTrue((withoutExt.OptionFlags & 0x0004) == 0x0000);

            // Something about continues...


            // Write out and re-read
            wb = WriteOutAndReadBack(wb);
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Check it's the same now
            withExt = wb.Workbook.GetSSTString(0);
            withoutExt = wb.Workbook.GetSSTString(31);

            Assert.AreEqual("O:Alloc:Qty", withExt.String);
            Assert.IsTrue((withExt.OptionFlags & 0x0004) == 0x0004);

            Assert.AreEqual("RT", withoutExt.String);
            Assert.IsTrue((withoutExt.OptionFlags & 0x0004) == 0x0000);
        }
        [Test]
        public void Test48026()
        {
            OpenSample("48026.xls");
        }
        [Test]
        public void Test48968()
        {
            HSSFWorkbook wb = OpenSample("48968.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);

            DataFormatter fmt = new DataFormatter();

            // Check the dates
            ISheet s = wb.GetSheetAt(0);
            ICell cell_d20110325 = s.GetRow(0).GetCell(0);
            ICell cell_d19000102 = s.GetRow(11).GetCell(0);
            ICell cell_d19000100 = s.GetRow(21).GetCell(0);
            Assert.AreEqual(s.GetRow(0).GetCell(3).StringCellValue, fmt.FormatCellValue(cell_d20110325));
            Assert.AreEqual(s.GetRow(11).GetCell(3).StringCellValue, fmt.FormatCellValue(cell_d19000102));
            // There is no such thing as 00/01/1900...
            Assert.AreEqual("00/01/1900 06:14:24", s.GetRow(21).GetCell(3).StringCellValue);
            Assert.AreEqual("31/12/1899 06:14:24", fmt.FormatCellValue(cell_d19000100));

            // Check the cached values
            Assert.AreEqual("HOUR(A1)", s.GetRow(5).GetCell(0).CellFormula);
            Assert.AreEqual(11.0, s.GetRow(5).GetCell(0).NumericCellValue);
            Assert.AreEqual("MINUTE(A1)", s.GetRow(6).GetCell(0).CellFormula);
            Assert.AreEqual(39.0, s.GetRow(6).GetCell(0).NumericCellValue);
            Assert.AreEqual("SECOND(A1)", s.GetRow(7).GetCell(0).CellFormula);
            Assert.AreEqual(54.0, s.GetRow(7).GetCell(0).NumericCellValue);

            // Re-evaulate and check
            HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
            Assert.AreEqual("HOUR(A1)", s.GetRow(5).GetCell(0).CellFormula);
            Assert.AreEqual(11.0, s.GetRow(5).GetCell(0).NumericCellValue);
            Assert.AreEqual("MINUTE(A1)", s.GetRow(6).GetCell(0).CellFormula);
            Assert.AreEqual(39.0, s.GetRow(6).GetCell(0).NumericCellValue);
            Assert.AreEqual("SECOND(A1)", s.GetRow(7).GetCell(0).CellFormula);
            Assert.AreEqual(54.0, s.GetRow(7).GetCell(0).NumericCellValue);

            // Push the time forward a bit and check
            double date = s.GetRow(0).GetCell(0).NumericCellValue;
            s.GetRow(0).GetCell(0).SetCellValue(date + 1.26);

            HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
            Assert.AreEqual("HOUR(A1)", s.GetRow(5).GetCell(0).CellFormula);
            Assert.AreEqual(11.0 + 6.0, s.GetRow(5).GetCell(0).NumericCellValue);
            Assert.AreEqual("MINUTE(A1)", s.GetRow(6).GetCell(0).CellFormula);
            Assert.AreEqual(39.0 + 14.0 + 1, s.GetRow(6).GetCell(0).NumericCellValue);
            Assert.AreEqual("SECOND(A1)", s.GetRow(7).GetCell(0).CellFormula);
            Assert.AreEqual(54.0 + 24.0 - 60, s.GetRow(7).GetCell(0).NumericCellValue);
        }
        /**
        * Problem with cloning a sheet with a chart
        *  contained in it.
        */
        [Test]
        public void Test49096()
        {
            HSSFWorkbook wb = OpenSample("49096.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);

            Assert.IsNotNull(wb.GetSheetAt(0));
            wb.CloneSheet(0);
            Assert.AreEqual(2, wb.NumberOfSheets);

            wb = WriteOutAndReadBack(wb);
            Assert.AreEqual(2, wb.NumberOfSheets);
        }

        [Test]
        public void Test49219()
        {
            HSSFWorkbook wb = OpenSample("49219.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);
            Assert.AreEqual("DGATE", wb.GetSheetAt(0).GetRow(1).GetCell(0).StringCellValue);
        }
        /**
         * Setting the user style name on custom styles
         */
        [Test]
        public void Test49689()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            ISheet s = wb1.CreateSheet("Test");
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);

            HSSFCellStyle cs1 = (HSSFCellStyle)wb1.CreateCellStyle();
            HSSFCellStyle cs2 = (HSSFCellStyle)wb1.CreateCellStyle();
            HSSFCellStyle cs3 = (HSSFCellStyle)wb1.CreateCellStyle();

            Assert.AreEqual(21, cs1.Index);
            cs1.UserStyleName = ("Testing");

            Assert.AreEqual(22, cs2.Index);
            cs2.UserStyleName = ("Testing 2");

            Assert.AreEqual(23, cs3.Index);
            cs3.UserStyleName = ("Testing 3");

            // Set one
            c.CellStyle = (cs1);

            // Write out and read back
            HSSFWorkbook wb2 = WriteOutAndReadBack(wb1);

            // Re-check
            Assert.AreEqual("Testing", ((HSSFCellStyle)wb2.GetCellStyleAt((short)21)).UserStyleName);
            Assert.AreEqual("Testing 2", ((HSSFCellStyle)wb2.GetCellStyleAt((short)22)).UserStyleName);
            Assert.AreEqual("Testing 3", ((HSSFCellStyle)wb2.GetCellStyleAt((short)23)).UserStyleName);

            wb2.Close();
        }
        [Test]
        public void Test49751()
        {
            HSSFWorkbook wb = OpenSample("49751.xls");
            int numCellStyles = wb.NumCellStyles;
            string[] namedStyles = new string[]{
            "20% - Accent1", "20% - Accent2", "20% - Accent3", "20% - Accent4", "20% - Accent5",
            "20% - Accent6", "40% - Accent1", "40% - Accent2", "40% - Accent3", "40% - Accent4",
            "40% - Accent5", "40% - Accent6", "60% - Accent1", "60% - Accent2", "60% - Accent3",
            "60% - Accent4", "60% - Accent5", "60% - Accent6", "Accent1", "Accent2", "Accent3",
            "Accent4", "Accent5", "Accent6", "Bad", "Calculation", "Check Cell", "Explanatory Text",
            "Good", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Input", "Linked Cell",
            "Neutral", "Note", "Output", "Title", "Total", "Warning Text"};

            IList<string> namedStylesList = Arrays.AsList(namedStyles);

            List<String> collecteddStyles = new List<String>();
            for (int i = 0; i < numCellStyles; i++)
            {
                HSSFCellStyle cellStyle = (HSSFCellStyle)wb.GetCellStyleAt(i);
                String styleName = cellStyle.UserStyleName;
                if (styleName != null)
                {
                    collecteddStyles.Add(styleName);
                    Assert.IsTrue(namedStylesList.Contains(styleName));
                }
            }
        }
        /**
        * Last row number when shifting rows
        */
        [Test]
        public void Test50416LastRowNumber()
        {
            // Create the workbook with 1 sheet which contains 3 rows
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Bug50416");
            IRow row1 = sheet.CreateRow(0);
            ICell cellA_1 = row1.CreateCell(0, CellType.String);
            cellA_1.SetCellValue("Cell A,1");
            IRow row2 = sheet.CreateRow(1);
            ICell cellA_2 = row2.CreateCell(0, CellType.String);
            cellA_2.SetCellValue("Cell A,2");
            IRow row3 = sheet.CreateRow(2);
            ICell cellA_3 = row3.CreateCell(0, CellType.String);
            cellA_3.SetCellValue("Cell A,3");

            // Test the last Row number it currently correct
            Assert.AreEqual(2, sheet.LastRowNum);

            // Shift the first row to the end
            sheet.ShiftRows(0, 0, 3);
            Assert.AreEqual(3, sheet.LastRowNum);
            Assert.AreEqual(-1, sheet.GetRow(0).LastCellNum);
            Assert.AreEqual("Cell A,2", sheet.GetRow(1).GetCell(0).StringCellValue);
            Assert.AreEqual("Cell A,3", sheet.GetRow(2).GetCell(0).StringCellValue);
            Assert.AreEqual("Cell A,1", sheet.GetRow(3).GetCell(0).StringCellValue);

            // Shift the 2nd row up to the first one
            sheet.ShiftRows(1, 1, -1);
            Assert.AreEqual(3, sheet.LastRowNum);
            Assert.AreEqual("Cell A,2", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual(-1, sheet.GetRow(1).LastCellNum);
            Assert.AreEqual("Cell A,3", sheet.GetRow(2).GetCell(0).StringCellValue);
            Assert.AreEqual("Cell A,1", sheet.GetRow(3).GetCell(0).StringCellValue);

            // Shift the 4th row up into the gap in the 3rd row
            sheet.ShiftRows(3, 3, -2);
            Assert.AreEqual(2, sheet.LastRowNum);
            Assert.AreEqual("Cell A,2", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("Cell A,1", sheet.GetRow(1).GetCell(0).StringCellValue);
            Assert.AreEqual("Cell A,3", sheet.GetRow(2).GetCell(0).StringCellValue);
            Assert.AreEqual(-1, sheet.GetRow(3).LastCellNum);

            // Now zap the empty 4th row - won't do anything
            sheet.RemoveRow(sheet.GetRow(3));

            // Test again the last row number which should be 2
            Assert.AreEqual(2, sheet.LastRowNum);
            Assert.AreEqual("Cell A,2", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("Cell A,1", sheet.GetRow(1).GetCell(0).StringCellValue);
            Assert.AreEqual("Cell A,3", sheet.GetRow(2).GetCell(0).StringCellValue);

            workbook.Close();
        }
        [Test]
        public void Test50426()
        {
            HSSFWorkbook wb = OpenSample("50426.xls");
            WriteOutAndReadBack(wb);
        }
        /**
         * If you send a file between Excel and OpenOffice enough, something
         *  will turn the "General" format into "GENERAL"
         */
        [Test]
        public void Test50756()
        {
            HSSFWorkbook wb = OpenSample("50756.xls");
            ISheet s = wb.GetSheetAt(0);
            IRow r17 = s.GetRow(16);
            IRow r18 = s.GetRow(17);
            DataFormatter df = new DataFormatter();

            Assert.AreEqual(10.0, r17.GetCell(1).NumericCellValue);
            Assert.AreEqual(20.0, r17.GetCell(2).NumericCellValue);
            Assert.AreEqual(20.0, r17.GetCell(3).NumericCellValue);
            Assert.AreEqual("GENERAL", r17.GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("GENERAL", r17.GetCell(2).CellStyle.GetDataFormatString());
            Assert.AreEqual("GENERAL", r17.GetCell(3).CellStyle.GetDataFormatString());
            Assert.AreEqual("10", df.FormatCellValue(r17.GetCell(1)));
            Assert.AreEqual("20", df.FormatCellValue(r17.GetCell(2)));
            Assert.AreEqual("20", df.FormatCellValue(r17.GetCell(3)));

            Assert.AreEqual(16.0, r18.GetCell(1).NumericCellValue);
            Assert.AreEqual(35.0, r18.GetCell(2).NumericCellValue);
            Assert.AreEqual(123.0, r18.GetCell(3).NumericCellValue);
            Assert.AreEqual("GENERAL", r18.GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("GENERAL", r18.GetCell(2).CellStyle.GetDataFormatString());
            Assert.AreEqual("GENERAL", r18.GetCell(3).CellStyle.GetDataFormatString());
            Assert.AreEqual("16", df.FormatCellValue(r18.GetCell(1)));
            Assert.AreEqual("35", df.FormatCellValue(r18.GetCell(2)));
            Assert.AreEqual("123", df.FormatCellValue(r18.GetCell(3)));
        }
        public void Test50779()
        {
            HSSFWorkbook wb1 = OpenSample("50779_1.xls");
            WriteOutAndReadBack(wb1);

            HSSFWorkbook wb2 = OpenSample("50779_2.xls");
            WriteOutAndReadBack(wb2);
        }
        /**
         * A protected sheet with comments, when written out by
         *  POI, ends up upsetting excel.
         * TODO Identify the cause and add extra asserts for
         *  the bit excel cares about
         */
        [Test]
        public void Test50833()
        {
            HSSFWorkbook wb = OpenSample("50833.xls");
            ISheet s = wb.GetSheetAt(0);
            Assert.AreEqual("Sheet1", s.SheetName);
            Assert.AreEqual(false, s.Protect);

            ICell c = s.GetRow(0).GetCell(0);
            Assert.AreEqual("test cell value", c.RichStringCellValue.String);

            IComment cmt = c.CellComment;
            Assert.IsNotNull(cmt);
            Assert.AreEqual("Robert Lawrence", cmt.Author);
            Assert.AreEqual("Robert Lawrence:\ntest comment", cmt.String.String);

            // Reload
            wb = WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            c = s.GetRow(0).GetCell(0);

            // Re-check the comment
            cmt = c.CellComment;
            Assert.IsNotNull(cmt);
            Assert.AreEqual("Robert Lawrence", cmt.Author);
            Assert.AreEqual("Robert Lawrence:\ntest comment", cmt.String.String);

            // TODO Identify what excel doesn't like, and check for that
        }
        /**
         * The spec says that ChartEndObjectRecord has 6 reserved
         *  bytes on the end, but we sometimes find files without... 
         */
        [Test]
        public void Test50939()
        {
            HSSFWorkbook wb = OpenSample("50939.xls");
            Assert.AreEqual(2, wb.NumberOfSheets);
        }

        /**
         * File with exactly 256 data blocks (+header block)
         *  shouldn't break on POIFS loading 
         */
        [Test]
        public void Test51461()
        {
            byte[] data = HSSFITestDataProvider.Instance.GetTestDataFileContent("51461.xls");

            HSSFWorkbook wbPOIFS = new HSSFWorkbook(new POIFSFileSystem(
                  new MemoryStream(data)).Root, false);
            HSSFWorkbook wbNPOIFS = new HSSFWorkbook(new NPOIFSFileSystem(
                  new MemoryStream(data)).Root, false);

            Assert.AreEqual(2, wbPOIFS.NumberOfSheets);
            Assert.AreEqual(2, wbNPOIFS.NumberOfSheets);
        }
        /**
        * Large row numbers and NPOIFS vs POIFS
        */
        [Test]
        public void Test51535()
        {
            byte[] data = HSSFITestDataProvider.Instance.GetTestDataFileContent("51535.xls");

            HSSFWorkbook wbPOIFS = new HSSFWorkbook(new POIFSFileSystem(
                  new MemoryStream(data)).Root, false);
            HSSFWorkbook wbNPOIFS = new HSSFWorkbook(new NPOIFSFileSystem(
                  new MemoryStream(data)).Root, false);

            foreach (HSSFWorkbook wb in new HSSFWorkbook[] { wbPOIFS, wbNPOIFS })
            {
                Assert.AreEqual(3, wb.NumberOfSheets);

                // Check directly
                ISheet s = wb.GetSheetAt(0);
                Assert.AreEqual("Top Left Cell", s.GetRow(0).GetCell(0).StringCellValue);
                Assert.AreEqual("Top Right Cell", s.GetRow(0).GetCell(255).StringCellValue);
                Assert.AreEqual("Bottom Left Cell", s.GetRow(65535).GetCell(0).StringCellValue);
                Assert.AreEqual("Bottom Right Cell", s.GetRow(65535).GetCell(255).StringCellValue);

                // Extract and check
                ExcelExtractor ex = new ExcelExtractor(wb);
                String text = ex.Text;
                Assert.IsTrue(text.Contains("Top Left Cell"));
                Assert.IsTrue(text.Contains("Top Right Cell"));
                Assert.IsTrue(text.Contains("Bottom Left Cell"));
                Assert.IsTrue(text.Contains("Bottom Right Cell"));
            }
        }
        /**
     * Sum across multiple workbooks
     *  eg =SUM($Sheet2.A1:$Sheet3.A1)
     * DISABLED - We currently get the formula wrong, and mis-evaluate
     */
        public void DISABLEDtest48703()
        {
            HSSFWorkbook wb = OpenSample("48703.xls");
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Check reading the formula
            ISheet sheet = wb.GetSheetAt(0);
            IRow r = sheet.GetRow(0);
            ICell c = r.GetCell(0);

            Assert.AreEqual("SUM(Sheet2!A1:Sheet3!A1)", c.CellFormula);
            Assert.AreEqual(4.0, c.NumericCellValue);

            // Check the evaluated result
            HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);
            eval.EvaluateFormulaCell(c);
            Assert.AreEqual(4.0, c.NumericCellValue);
        }
        /**
         * Normally encrypted files have BOF then FILEPASS, but
         *  some may squeeze a WRITEPROTECT in the middle
         */
        [Test]
        public void Test51832()
        {
            try
            {
                OpenSample("51832.xls");
                Assert.Fail("Encrypted file");
            }
            catch (EncryptedDocumentException)
            {
                // Good
            }
        }
        [Test]
        public void Test49896()
        {
            HSSFWorkbook wb = OpenSample("49896.xls");
            HSSFCell cell = (HSSFCell)wb.GetSheetAt(0).GetRow(1).GetCell(1);
            char separator = Path.DirectorySeparatorChar;
            Assert.AreEqual("VLOOKUP(A2,'[C:Documents and Settings" + separator + "Yegor" + separator
                + "My Documents" + separator + "csco.xls]Sheet1'!$A$2:$B$3,2,FALSE)",
                    cell.CellFormula);
        }
        [Test]
        public void Test49529()
        {
            // user code reported in Bugzilla #49529
            HSSFWorkbook workbook = OpenSample("49529.xls");
            workbook.GetSheetAt(0).CreateDrawingPatriarch();
            // prior to the fix the line below failed with
            // java.lang.InvalidOperationException: EOF - next record not available
            workbook.CloneSheet(0);

            // make sure we are still readable
            WriteOutAndReadBack(workbook);
        }

        [Test]
        public void Test51670()
        {
            HSSFWorkbook wb = OpenSample("51670.xls");
            WriteOutAndReadBack(wb);
        }
        /**
         * Note - part of this test is still failing, see
         * {@link TestUnfixedBugs#test49612()}
         */
        [Test]
        public void bug49612_part()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("49612.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            HSSFRow row = sh.GetRow(0) as HSSFRow;
            HSSFCell c1 = row.GetCell(2) as HSSFCell;
            HSSFCell d1 = row.GetCell(3) as HSSFCell;
            HSSFCell e1 = row.GetCell(2) as HSSFCell;

            Assert.AreEqual("SUM(BOB+JIM)", c1.CellFormula);

            // Problem 1: See TestUnfixedBugs#test49612()
            // Problem 2: TestUnfixedBugs#test49612()

            // Problem 3: These used to fail, now pass
            HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);
            Assert.AreEqual(30.0, eval.Evaluate(c1).NumberValue, 0.001, "Evaluating c1");
            Assert.AreEqual(30.0, eval.Evaluate(d1).NumberValue, 0.001, "Evaluating d1");
            Assert.AreEqual(30.0, eval.Evaluate(e1).NumberValue, 0.001, "Evaluating e1");
        }

        [Test]
        public void bug51675()
        {
            List<short> list = new List<short>();
            HSSFWorkbook workbook = OpenSample("51675.xls");
            HSSFSheet sh = workbook.GetSheetAt(0) as HSSFSheet;
            InternalSheet ish = HSSFTestHelper.GetSheetForTest(sh);
            PageSettingsBlock psb = (PageSettingsBlock)ish.Records[(13)];
            psb.VisitContainedRecords(new RecordVisitor1(list));
            Assert.IsTrue(list[(list.Count - 1)] == UnknownRecord.BITMAP_00E9);
            Assert.IsTrue(list[(list.Count - 2)] == UnknownRecord.HEADER_FOOTER_089C);
        }
        public class RecordVisitor1 : RecordVisitor
        {
            private List<short> list;
            public RecordVisitor1(List<short> list)
            {
                this.list = list;
            }

            public void VisitRecord(NPOI.HSSF.Record.Record r)
            {
                list.Add(r.Sid);
            }
        }

        [Test]
        public void Test52272()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch p = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape s = p.CreateSimpleShape(new HSSFClientAnchor());
            s.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);

            HSSFSheet sh2 = wb.CloneSheet(0) as HSSFSheet;
            Assert.IsNotNull(sh2.DrawingPatriarch);

            wb.Close();
        }

        [Test]
        public void Test53432()
        {
            IWorkbook wb1 = new HSSFWorkbook(); //or new HSSFWorkbook();
            wb1.AddPicture(new byte[] { 123, 22 }, PictureType.JPEG);
            Assert.AreEqual(wb1.GetAllPictures().Count, 1);
            wb1.Close();

            wb1.Close();
            wb1 = new HSSFWorkbook();
            IWorkbook wb2 = WriteOutAndReadBack((HSSFWorkbook)wb1);
            wb1.Close();
            Assert.AreEqual(wb2.GetAllPictures().Count, 0);
            wb2.AddPicture(new byte[] { 123, 22 }, PictureType.JPEG);
            Assert.AreEqual(wb2.GetAllPictures().Count, 1);

            IWorkbook wb3 = WriteOutAndReadBack((HSSFWorkbook)wb2);
            wb2.Close();
            Assert.AreEqual(wb3.GetAllPictures().Count, 1);

            wb3.Close();
        }
        [Test]
        public void Test46250()
        {
            IWorkbook wb = OpenSample("46250.xls");
            ISheet sh = wb.GetSheet("Template");
            ISheet cSh = wb.CloneSheet(wb.GetSheetIndex(sh));

            HSSFPatriarch patriarch = (HSSFPatriarch)cSh.CreateDrawingPatriarch();
            HSSFTextbox tb = (HSSFTextbox)patriarch.Children[2];

            tb.String = (new HSSFRichTextString("POI test"));
            tb.Anchor = (new HSSFClientAnchor(0, 0, 0, 0, (short)0, 0, (short)10, 10));

            wb = WriteOutAndReadBack((HSSFWorkbook)wb);
        }

        [Test]
        public void Test53404()
        {
            IWorkbook wb = OpenSample("53404.xls");
            ISheet sheet = wb.GetSheet("test-sheet");
            int rowCount = sheet.LastRowNum + 1;
            int newRows = 5;
            for (int r = rowCount; r < rowCount + newRows; r++)
            {
                IRow row = sheet.CreateRow(r);
                row.CreateCell(0).SetCellValue(1.03 * (r + 7));
                row.CreateCell(1).SetCellValue(DateTime.Now);
                row.CreateCell(2).SetCellValue(DateTime.Now);
                row.CreateCell(3).SetCellValue(String.Format("row:{0}/col:{1}", r, 3));
                row.CreateCell(4).SetCellValue(true);
                row.CreateCell(5).SetCellType(CellType.Error);
                row.CreateCell(6).SetCellValue("added cells.");
            }

            wb = WriteOutAndReadBack((HSSFWorkbook)wb);
        }

        [Test]
        public void Test54016()
        {
            // This used to break
            HSSFWorkbook wb = OpenSample("54016.xls");
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }
        [Test]
        public void TestBug49237()
        {
            HSSFWorkbook wb = OpenSample("49237.xls");
            ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICellStyle rstyle = row.RowStyle;
            Assert.AreEqual(rstyle.BorderBottom, BorderStyle.Double);
        }
        [Test]
        public void Bug35897()
        {
            // password is abc
            try
            {
                Biff8EncryptionKey.CurrentUserPassword = ("abc");
                OpenSample("xor-encryption-abc.xls");
            }
            finally
            {
                Biff8EncryptionKey.CurrentUserPassword = (null);
            }

            // One using the only-recently-documented encryption header type 4,
            //  and the RC4 CryptoAPI encryption header structure
            try
            {
                OpenSample("35897-type4.xls");
                Assert.Fail("POI doesn't currently support the RC4 CryptoAPI encryption header structure");
            }
            catch (EncryptedDocumentException) { }
        }

        [Test]
        public void Bug56450()
        {
            HSSFWorkbook wb = OpenSample("56450.xls");
            HSSFSheet sheet = wb.GetSheetAt(0) as HSSFSheet;
            int comments = 0;
            foreach (IRow r in sheet)
            {
                foreach (ICell c in r)
                {
                    if (c.CellComment != null)
                    {
                        Assert.IsNotNull(c.CellComment.String.String);
                        comments++;
                    }
                }
            }
            Assert.AreEqual(0, comments);
        }

        [Test]
        public void Bug56482()
        {
            HSSFWorkbook wb = OpenSample("56482.xls");
            Assert.AreEqual(1, wb.NumberOfSheets);

            HSSFSheet sheet = wb.GetSheetAt(0) as HSSFSheet;
            HSSFSheetConditionalFormatting cf = sheet.SheetConditionalFormatting as HSSFSheetConditionalFormatting;

            Assert.AreEqual(5, cf.NumConditionalFormattings);
        }

        [Test]
        public void Bug56325()
        {
            HSSFWorkbook wb1;

            FileInfo file = HSSFTestDataSamples.GetSampleFile("56325.xls");
            Stream stream = new FileStream(file.FullName, FileMode.Open, FileAccess.ReadWrite);
            try
            {
                POIFSFileSystem fs = new POIFSFileSystem(stream);
                wb1 = new HSSFWorkbook(fs);
            }
            finally
            {
                stream.Close();
            }

            Assert.AreEqual(3, wb1.NumberOfSheets);
            wb1.RemoveSheetAt(0);
            Assert.AreEqual(2, wb1.NumberOfSheets);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            Assert.AreEqual(2, wb2.NumberOfSheets);
            wb2.RemoveSheetAt(0);
            Assert.AreEqual(1, wb2.NumberOfSheets);
            wb2.RemoveSheetAt(0);
            Assert.AreEqual(0, wb2.NumberOfSheets);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();

            Assert.AreEqual(0, wb3.NumberOfSheets);
            wb3.Close();
        }

        [Test]
        public void Bug56325a()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("56325a.xls");

            HSSFSheet sheet = wb1.CloneSheet(2) as HSSFSheet;
            wb1.SetSheetName(3, "Clone 1");
            sheet.RepeatingRows = (/*setter*/CellRangeAddress.ValueOf("2:3"));
            wb1.SetPrintArea(3, "$A$4:$C$10");

            sheet = wb1.CloneSheet(2) as HSSFSheet;
            wb1.SetSheetName(4, "Clone 2");
            sheet.RepeatingRows = (/*setter*/CellRangeAddress.ValueOf("2:3"));
            wb1.SetPrintArea(4, "$A$4:$C$10");

            wb1.RemoveSheetAt(2);

            IWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            Assert.AreEqual(4, wb2.NumberOfSheets);

            wb2.Close();
            wb1.Close();
        }


        /**
         * Formulas which reference named ranges, either in other
         *  sheets, or workbook scoped but in other workbooks.
         * Currently failing with 
         * java.lang.Exception: Unexpected eval class (NPOI.SS.Formula.Eval.NameXEval)
         */
        [Test]
        public void Bug56737()
        {
            IWorkbook wb = OpenSample("56737.xls");

            // Check the named range defInitions
            IName nSheetScope = wb.GetName("NR_To_A1");
            IName nWBScope = wb.GetName("NR_Global_B2");

            Assert.IsNotNull(nSheetScope);
            Assert.IsNotNull(nWBScope);

            Assert.AreEqual("Defines!$A$1", nSheetScope.RefersToFormula);
            Assert.AreEqual("Defines!$B$2", nWBScope.RefersToFormula);

            // Check the different kinds of formulas
            ISheet s = wb.GetSheetAt(0);
            ICell cRefSName = s.GetRow(1).GetCell(3);
            ICell cRefWName = s.GetRow(2).GetCell(3);

            Assert.AreEqual("Defines!NR_To_A1", cRefSName.CellFormula);
            // TODO How does Excel know to prefix this with the filename?
            // This is what Excel itself shows
            //assertEquals("'56737.xls'!NR_Global_B2", cRefWName.getCellFormula());
            // TODO This isn't right, but it's what we currently generate....
            Assert.AreEqual("NR_Global_B2", cRefWName.CellFormula);


            // Try to Evaluate them
            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual("Test A1", eval.Evaluate(cRefSName).StringValue);
            Assert.AreEqual(142, (int)eval.Evaluate(cRefWName).NumberValue);

            // Try to Evaluate everything
            eval.EvaluateAll();

        }

        /**
         * InvalidCastException in HSSFOptimiser - StyleRecord cannot be cast to 
         * ExtendedFormatRecord when removing un-used styles
         */
        [Test]
        public void Bug54443()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFCellStyle style = workbook.CreateCellStyle() as HSSFCellStyle;
            HSSFCellStyle newStyle = workbook.CreateCellStyle() as HSSFCellStyle;

            HSSFSheet mySheet = workbook.CreateSheet() as HSSFSheet;
            HSSFRow row = mySheet.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;

            // Use style
            cell.CellStyle = (/*setter*/style);
            // Switch to newStyle, style is now un-used
            cell.CellStyle = (/*setter*/newStyle);

            // Optimise
            HSSFOptimiser.OptimiseCellStyles(workbook);
        }

        /**
         * Intersection formula ranges, eg =(C2:D3 D3:E4)
         */
        [Test]
        public void Bug52111()
        {
            IWorkbook wb = OpenSample("Intersection-52111.xls");
            ISheet s = wb.GetSheetAt(0);
            assertFormula(wb, s.GetRow(2).GetCell(0), "(C2:D3 D3:E4)", "4");
            assertFormula(wb, s.GetRow(6).GetCell(0), "Tabelle2!E:E Tabelle2!$A11:$IV11", "5");
            assertFormula(wb, s.GetRow(8).GetCell(0), "Tabelle2!E:F Tabelle2!$A11:$IV12", null);
        }

        private void assertFormula(IWorkbook wb, ICell intF, String expectedFormula, String expectedResultOrNull)
        {
            Assert.AreEqual(CellType.Formula, intF.CellType);
            if (null == expectedResultOrNull)
            {
                Assert.AreEqual(CellType.Error, intF.CachedFormulaResultType);
                expectedResultOrNull = "#VALUE!";
            }
            else
            {
                Assert.AreEqual(CellType.Numeric, intF.CachedFormulaResultType);
            }

            Assert.AreEqual(expectedFormula, intF.CellFormula);

            // Check we can evaluate it correctly
            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual(expectedResultOrNull, eval.Evaluate(intF).FormatAsString());
        }
        [Test]
        public void Bug42016()
        {
            IWorkbook wb = OpenSample("42016.xls");
            ISheet s = wb.GetSheetAt(0);
            for (int row = 0; row < 7; row++)
            {
                Assert.AreEqual("A$1+B$1", s.GetRow(row).GetCell(2).CellFormula);
            }
        }

        /**
         * Unexpected record type (NPOI.HSSF.Record.ColumnInfoRecord)
         */
        [Test]
        public void Bug53984()
        {
            IWorkbook wb = OpenSample("53984.xls");
            ISheet s = wb.GetSheetAt(0);
            Assert.AreEqual("International Communication Services SA", s.GetRow(2).GetCell(0).StringCellValue);
            Assert.AreEqual("Saudi Arabia-Riyadh", s.GetRow(210).GetCell(0).StringCellValue);
        }

        /**
     * Read, Write, read for formulas point to cells in other files.
     * See {@link #bug46670()} for the main test, this just
     *  covers Reading an existing file and Checking it.
     * TODO Fix this so that it works - formulas are ending up as
     *  #REF when being Changed
     */
        //    [Test]
        public void bug46670_existing()
        {
            HSSFWorkbook wb;
            ISheet s;
            ICell c;

            // Expected values
            String refLocal = "'[refs/airport.xls]Sheet1'!$A$2";
            String refHttp = "'[9http://www.principlesofeconometrics.com/excel/airline.xls]Sheet1'!$A$2";

            // Check we can read them correctly
            wb = OpenSample("46670_local.xls");
            s = wb.GetSheetAt(0);
            Assert.AreEqual(refLocal, s.GetRow(0).GetCell(0).CellFormula);

            wb = OpenSample("46670_http.xls");
            s = wb.GetSheetAt(0);
            Assert.AreEqual(refHttp, s.GetRow(0).GetCell(0).CellFormula);

            // Now try to Set them to the same values, and ensure that
            //  they end up as they did before, even with a save and re-load
            wb = OpenSample("46670_local.xls");
            s = wb.GetSheetAt(0);
            c = s.GetRow(0).GetCell(0);
            c.CellFormula = (/*setter*/refLocal);
            Assert.AreEqual(refLocal, c.CellFormula);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            Assert.AreEqual(refLocal, s.GetRow(0).GetCell(0).CellFormula);


            wb = OpenSample("46670_http.xls");
            s = wb.GetSheetAt(0);
            c = s.GetRow(0).GetCell(0);
            c.CellFormula = (/*setter*/refHttp);
            Assert.AreEqual(refHttp, c.CellFormula);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            Assert.AreEqual(refHttp, s.GetRow(0).GetCell(0).CellFormula);
        }

        [Test]
        public void Test57163()
        {
            IWorkbook wb = OpenSample("57163.xls");

            while (wb.NumberOfSheets > 1)
            {
                wb.RemoveSheetAt(1);
            }
            wb.Close();
        }

        [Test]
        public void Test53109()
        {
            HSSFWorkbook wb = OpenSample("53109.xls");

            IWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);

            wb.Close();
        }

        [Test]
        public void Test53109a()
        {
            HSSFWorkbook wb = OpenSample("com.aida-tour.www_SPO_files_maldives%20august%20october.xls");

            IWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);

            wb.Close();
        }
        [Test]
        public void Test48043()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("56325a.xls");

            wb.RemoveSheetAt(2);
            wb.RemoveSheetAt(1);

            //Sheet s = wb.createSheet("sheetname");
            ISheet s = wb.GetSheetAt(0);
            IRow row = s.CreateRow(0);
            ICell cell = row.CreateCell(0);

            cell.SetCellFormula(
                    "IF(AND(ISBLANK(A10)," +
                    "ISBLANK(B10)),\"\"," +
                    "CONCATENATE(A10,\"-\",B10))");

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            eval.EvaluateAll();

            /*OutputStream out = new FileOutputStream("C:\\temp\\48043.xls");
            try {
              wb.write(out);
            } finally {
              out.close();
            }*/

            IWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);
            wbBack.Close();

            wb.Close();
        }
        [Test]
        public void Test57925()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("57925.xls");

            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row)
                    {
                        new DataFormatter().FormatCellValue(cell);
                    }
                }
            }

            wb.Close();
        }


        [Test]
        public void Test46515()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("46515.xls");
            // Get structure from webservice
            String urlString = "http://poi.apache.org/resources/images/project-logo.jpg";
            Uri structURL = new Uri(urlString);
            Image bimage;
            try
            {
                WebClient client = new WebClient();
                MemoryStream ms = new MemoryStream(client.DownloadData(structURL));
                bimage = Image.Load(ms);
                ms.Close();
            }
            catch (WebException)
            {
                //Assume.assumeNoException("Downloading a jpg from poi.apache.org should work", e);
                return;
            }
            // Convert BufferedImage to byte[]
            ByteArrayOutputStream imageBAOS = new ByteArrayOutputStream();
            //ImageIO.write(bimage, "jpeg", imageBAOS);
            bimage.SaveAsJpeg(imageBAOS);
            imageBAOS.Flush();
            byte[] imageBytes = imageBAOS.ToByteArray();
            imageBAOS.Close();
            // Pop structure into Structure HSSFSheet
            int pict = wb.AddPicture(imageBytes, PictureType.JPEG);
            ISheet sheet = wb.GetSheet("Structure");
            Assert.IsNotNull(sheet, "Did not find sheet");
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)1, 1, (short)10, 22);
            anchor.AnchorType = AnchorType.MoveDontResize; //(2);
            patriarch.CreatePicture(anchor, pict);
            // Write out destination file
            //        FileOutputStream fileOut = new FileOutputStream("/tmp/46515.xls");
            //        wb.write(fileOut);
            //        fileOut.close();
            wb.Close();
        }
        [Test]
        public void Test55668()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("55668.xls");
            ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("IF(TRUE,\"\",\"\")", cell.CellFormula);
            Assert.AreEqual("", cell.StringCellValue);
            cell.SetCellType(CellType.String);
            Assert.AreEqual(CellType.Blank, cell.CellType);
            try
            {
                Assert.IsNull(cell.CellFormula);
                Assert.Fail("Should throw an exception here");
            }
            catch (InvalidOperationException e)
            {
                // expected here
            }
            Assert.AreEqual("", cell.StringCellValue);
            wb.Close();
        }

        [Test]
        public void Test55982()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("55982.xls");
            ISheet newSheet = wb.CloneSheet(1);
            Assert.IsNotNull(newSheet);
            wb.Close();
        }

        /**
         * Test generator of ids for the CommonObjectDataSubRecord record.
         */
        [Test]
        public void Test51332()
        {
            HSSFClientAnchor anchor = new HSSFClientAnchor();
            HSSFSimpleShape shape;
            CommonObjectDataSubRecord cmo;

            shape = new HSSFTextbox(null, anchor);
            shape.ShapeId = 1025;
            cmo = (CommonObjectDataSubRecord)shape.GetObjRecord().SubRecords[0];
            Assert.AreEqual(1, cmo.ObjectId);
            shape = new HSSFPicture(null, anchor);
            shape.ShapeId = 1026;
            cmo = (CommonObjectDataSubRecord)shape.GetObjRecord().SubRecords[0];
            Assert.AreEqual(2, cmo.ObjectId);
            shape = new HSSFComment(null, anchor);
            shape.ShapeId = 1027;
            cmo = (CommonObjectDataSubRecord)shape.GetObjRecord().SubRecords[0];
            Assert.AreEqual(1027, cmo.ObjectId);
        }


        // As of POI 3.15 beta 2, LibreOffice does not display the diagonal border while it does display the bottom border
        // I have not checked Excel to know if this is a LibreOffice or a POI problem.
        [Test]
        public void Test53564()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet("Page 1") as HSSFSheet;
            short BLUE = 30;

            HSSFSheetConditionalFormatting scf = sheet.SheetConditionalFormatting as HSSFSheetConditionalFormatting;
            HSSFConditionalFormattingRule rule = scf.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan, "10") as HSSFConditionalFormattingRule;

            HSSFBorderFormatting bord = rule.CreateBorderFormatting() as HSSFBorderFormatting;
            bord.BorderDiagonal = BorderStyle.Thick;
            Assert.AreEqual(BorderStyle.Thick, bord.BorderDiagonal);
            bord.IsBackwardDiagonalOn = true;
            Assert.IsTrue(bord.IsBackwardDiagonalOn);
            bord.IsForwardDiagonalOn = true;
            Assert.IsTrue(bord.IsForwardDiagonalOn);

            bord.DiagonalBorderColor = BLUE;
            Assert.AreEqual(BLUE, bord.DiagonalBorderColor);
            // Create the bottom border style so we know what a border is supposed to look like
            bord.BorderBottom = BorderStyle.Thick;
            Assert.AreEqual(BorderStyle.Thick, bord.BorderBottom);
            bord.BottomBorderColor = BLUE;
            Assert.AreEqual(BLUE, bord.BottomBorderColor);

            CellRangeAddress[] A2_D4 = { new CellRangeAddress(1, 3, 0, 3) };
            scf.AddConditionalFormatting(A2_D4, rule);

            // Set a cell value within the conditional formatting range whose rule would resolve to True.
            ICell C3 = sheet.CreateRow(2).CreateCell(2);
            C3.SetCellValue(30.0);

            // Manually check the output file with Excel to see if the diagonal border is present
            //OutputStream fos = new FileOutputStream("/tmp/53564.xls");
            //wb.write(fos);
            //fos.Close();
            wb.Close();
        }


        // follow https://svn.apache.org/viewvc?view=revision&revision=1896552 to write a unit test for this fix.
        [Test]
        public void Test52447()
        {
            IWorkbook wb=null;
            try
            {
                wb = HSSFTestDataSamples.OpenSampleWorkbook("52447.xls");
            } catch { 
                Assert.IsNotNull(wb);
            }
        }

        [Test]
        public void TestAmazonDownloadFile()
        {
            IWorkbook wb = null;
            try
            {
                wb = HSSFTestDataSamples.OpenSampleWorkbook("ca_apparel_browse_tree_guide._TTH_.xls");
            }
            catch
            {
                Assert.IsNotNull(wb);
            }
        }
    }
}
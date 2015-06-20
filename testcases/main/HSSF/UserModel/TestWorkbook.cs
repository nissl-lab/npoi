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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Util;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NPOI.HSSF.Record.Aggregates;
    using TestCases.HSSF;

    using NUnit.Framework;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    /**
     * Class to Test Workbook functionality
     *
     * @author Andrew C. Oliver
     * @author Greg Merrill
     * @author Siggi Cherem
     */
    [TestFixture]
    public class TestWorkbook
    {
        private static String LAST_NAME_KEY = "lastName";
        private static String FIRST_NAME_KEY = "firstName";
        private static String SSN_KEY = "ssn";
        private static String REPLACE_ME = "replaceMe";
        private static String REPLACED = "replaced";
        private static String DO_NOT_REPLACE = "doNotReplace";
        private static String EMPLOYEE_INFORMATION = "Employee Info";
        private static String LAST_NAME_VALUE = "Bush";
        private static String FIRST_NAME_VALUE = "George";
        private static String SSN_VALUE = "555555555";
        private SanityChecker sanityChecker = new SanityChecker();


        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        /**
         * TEST NAME:  Test Write Sheet Simple <P>
         * OBJECTIVE:  Test that HSSF can Create a simple spreadsheet with numeric and string values.<P>
         * SUCCESS:    HSSF Creates a sheet.  Filesize Matches a known good.  NPOI.SS.UserModel.Sheet objects
         *             Last row, first row is Tested against the correct values (99,0).<P>
         * FAILURE:    HSSF does not Create a sheet or excepts.  Filesize does not Match the known good.
         *             NPOI.SS.UserModel.Sheet last row or first row is incorrect.             <P>
         *
         */
        [Test]
        public void TestWriteSheetSimple()
        {
            string filepath = TempFile.GetTempFilePath("TestWriteSheetSimple",
                                                        ".xls");
            FileStream out1 = new FileStream(filepath, FileMode.OpenOrCreate);
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            IRow r = null;
            ICell c = null;

            for (int rownum = 0; rownum < 100; rownum++)
            {
                r = s.CreateRow(rownum);

                for (int cellnum = 0; cellnum < 50; cellnum += 2)
                {
                    c = r.CreateCell(cellnum);
                    c.SetCellValue(rownum * 10000 + cellnum
                                   + (((double)rownum / 1000)
                                      + ((double)cellnum / 10000)));
                    c = r.CreateCell(cellnum + 1);
                    c.SetCellValue(new HSSFRichTextString("TEST"));
                }
            }
            wb.Write(out1);
            out1.Close();
            sanityChecker.CheckHSSFWorkbook(wb);
            Assert.AreEqual(99, s.LastRowNum, "LAST ROW == 99");
            Assert.AreEqual(0, s.FirstRowNum, "FIRST ROW == 0");
        }

        /**
         * TEST NAME:  Test Write/Modify Sheet Simple <P>
         * OBJECTIVE:  Test that HSSF can Create a simple spreadsheet with numeric and string values,
         *             Remove some rows, yet still have a valid file/data.<P>
         * SUCCESS:    HSSF Creates a sheet.  Filesize Matches a known good.  NPOI.SS.UserModel.Sheet objects
         *             Last row, first row is Tested against the correct values (74,25).<P>
         * FAILURE:    HSSF does not Create a sheet or excepts.  Filesize does not Match the known good.
         *             NPOI.SS.UserModel.Sheet last row or first row is incorrect.             <P>
         *
         */
        [Test]
        public void TestWriteModifySheetSimple()
        {
            string filepath = TempFile.GetTempFilePath("TestWriteSheetSimple",
                                                        ".xls");
            FileStream out1 = new FileStream(filepath, FileMode.OpenOrCreate);
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            IRow r = null;
            ICell c = null;

            for (int rownum = 0; rownum < 100; rownum++)
            {
                r = s.CreateRow(rownum);

                for (int cellnum = 0; cellnum < 50; cellnum += 2)
                {
                    c = r.CreateCell(cellnum);
                    c.SetCellValue(rownum * 10000 + cellnum
                                   + (((double)rownum / 1000)
                                      + ((double)cellnum / 10000)));
                    c = r.CreateCell(cellnum + 1);
                    c.SetCellValue(new HSSFRichTextString("TEST"));
                }
            }
            for (int rownum = 0; rownum < 25; rownum++)
            {
                r = s.GetRow(rownum);
                s.RemoveRow(r);
            }
            for (int rownum = 75; rownum < 100; rownum++)
            {
                r = s.GetRow(rownum);
                s.RemoveRow(r);
            }
            wb.Write(out1);
            out1.Close();

            sanityChecker.CheckHSSFWorkbook(wb);
            Assert.AreEqual(74, s.LastRowNum, "LAST ROW == 74");
            Assert.AreEqual(25, s.FirstRowNum, "FIRST ROW == 25");
        }

        /**
         * TEST NAME:  Test Read Simple <P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet (Simple.xls).<P>
         * SUCCESS:    HSSF reads the sheet.  Matches values in their particular positions.<P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF cannot identify values
         *             in the sheet in their known positions.<P>
         *
         */
        [Test]
        public void TestReadSimple()
        {
            HSSFWorkbook workbook = OpenSample("Simple.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);

            ICell cell = sheet.GetRow(0).GetCell(0);
            Assert.AreEqual(REPLACE_ME, cell.RichStringCellValue.String);
        }

        /**
         * TEST NAME:  Test Read Simple w/ Data Format<P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet (SimpleWithDataFormat.xls).<P>
         * SUCCESS:    HSSF reads the sheet.  Matches values in their particular positions and format is correct<P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF cannot identify values
         *             in the sheet in their known positions.<P>
         *
         */
        [Test]
        public void TestReadSimpleWithDataFormat()
        {
            HSSFWorkbook workbook = OpenSample("SimpleWithDataFormat.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            IDataFormat format = workbook.CreateDataFormat();
            ICell cell = sheet.GetRow(0).GetCell(0);

            Assert.AreEqual(1.25, cell.NumericCellValue, 1e-10);

            Assert.AreEqual(format.GetFormat(cell.CellStyle.DataFormat), "0.0");
        }

        /**
             * TEST NAME:  Test Read/Write Simple w/ Data Format<P>
             * OBJECTIVE:  Test that HSSF can Write a sheet with custom data formats and then read it and get the proper formats.<P>
             * SUCCESS:    HSSF reads the sheet.  Matches values in their particular positions and format is correct<P>
             * FAILURE:    HSSF does not read a sheet or excepts.  HSSF cannot identify values
             *             in the sheet in their known positions.<P>
             *
             */
        [Test]
        public void TestWriteDataFormat()
        {
            string filepath = TempFile.GetTempFilePath("TestWriteSheetSimple", ".xls");
            FileStream out1 = new FileStream(filepath, FileMode.OpenOrCreate);
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            IRow r = null;
            ICell c = null;
            IDataFormat format = wb.CreateDataFormat();
            NPOI.SS.UserModel.ICellStyle cs = wb.CreateCellStyle();

            short df = format.GetFormat("0.0");
            cs.DataFormat = (df);

            r = s.CreateRow(0);
            c = r.CreateCell(0);
            c.CellStyle = (cs);
            c.SetCellValue(1.25);

            wb.Write(out1);
            out1.Close();

            FileStream stream = new FileStream(filepath, FileMode.OpenOrCreate);
            POIFSFileSystem fs = new POIFSFileSystem(stream);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            ICell cell =
                             sheet.GetRow(0).GetCell(0);
            format = workbook.CreateDataFormat();

            Assert.AreEqual(1.25, cell.NumericCellValue, 1e-10);

            Assert.AreEqual(format.GetFormat(df), "0.0");

            Assert.AreEqual(format, workbook.CreateDataFormat());

            stream.Close();
        }

        /**
         * TEST NAME:  Test Read Employee Simple <P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet (Employee.xls).<P>
         * SUCCESS:    HSSF reads the sheet.  Matches values in their particular positions.<P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF cannot identify values
         *             in the sheet in their known positions.<P>
         *
         */
        [Test]
        public void TestReadEmployeeSimple()
        {
            HSSFWorkbook workbook = OpenSample("Employee.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);

            Assert.AreEqual(EMPLOYEE_INFORMATION, sheet.GetRow(1).GetCell(1).RichStringCellValue.String);
            Assert.AreEqual(LAST_NAME_KEY, sheet.GetRow(3).GetCell(2).RichStringCellValue.String);
            Assert.AreEqual(FIRST_NAME_KEY, sheet.GetRow(4).GetCell(2).RichStringCellValue.String);
            Assert.AreEqual(SSN_KEY, sheet.GetRow(5).GetCell(2).RichStringCellValue.String);
        }

        /**
         * TEST NAME:  Test Modify Sheet Simple <P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet with a string value and replace
         *             it with another string value.<P>
         * SUCCESS:    HSSF reads a sheet.  HSSF replaces the cell value with another cell value. HSSF
         *             Writes the sheet out to another file.  HSSF reads the result and ensures the value
         *             has been properly replaced.    <P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF does not Write the sheet or excepts.
         *             HSSF does not re-read the sheet or excepts.  Upon re-reading the sheet the value
         *             is incorrect or has not been replaced. <P>
         *
         */
        [Test]
        public void TestModifySimple()
        {
            HSSFWorkbook workbook = OpenSample("Simple.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            ICell cell = sheet.GetRow(0).GetCell(0);

            cell.SetCellValue(new HSSFRichTextString(REPLACED));

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = workbook.GetSheetAt(0);
            cell = sheet.GetRow(0).GetCell(0);
            Assert.AreEqual(REPLACED, cell.RichStringCellValue.String);
        }

        /**
         * TEST NAME:  Test Modify Sheet Simple With Skipped cells<P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet with string values and replace
         *             them with other string values while not replacing other cells.<P>
         * SUCCESS:    HSSF reads a sheet.  HSSF replaces the cell value with another cell value. HSSF
         *             Writes the sheet out to another file.  HSSF reads the result and ensures the value
         *             has been properly replaced and unreplaced values are still unreplaced.    <P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF does not Write the sheet or excepts.
         *             HSSF does not re-read the sheet or excepts.  Upon re-reading the sheet the value
         *             is incorrect or has not been replaced or the incorrect cell has its value replaced
         *             or is incorrect. <P>
         *
         */
        [Test]
        public void TestModifySimpleWithSkip()
        {
            HSSFWorkbook workbook = OpenSample("SimpleWithSkip.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            ICell cell = sheet.GetRow(0).GetCell(1);

            cell.SetCellValue(new HSSFRichTextString(REPLACED));
            cell = sheet.GetRow(1).GetCell(0);
            cell.SetCellValue(new HSSFRichTextString(REPLACED));

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);

            sheet = workbook.GetSheetAt(0);
            cell = sheet.GetRow(0).GetCell(1);
            Assert.AreEqual(REPLACED, cell.RichStringCellValue.String);
            cell = sheet.GetRow(0).GetCell(0);
            Assert.AreEqual(DO_NOT_REPLACE, cell.RichStringCellValue.String);
            cell = sheet.GetRow(1).GetCell(0);
            Assert.AreEqual(REPLACED, cell.RichStringCellValue.String);
            cell = sheet.GetRow(1).GetCell(1);
            Assert.AreEqual(DO_NOT_REPLACE, cell.RichStringCellValue.String);
        }

        /**
         * TEST NAME:  Test Modify Sheet With Styling<P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet with string values and replace
         *             them with other string values despite any styling.  In this release of HSSF styling will
         *             probably be lost and is NOT Tested.<P>
         * SUCCESS:    HSSF reads a sheet.  HSSF replaces the cell values with other cell values. HSSF
         *             Writes the sheet out to another file.  HSSF reads the result and ensures the value
         *             has been properly replaced.    <P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF does not Write the sheet or excepts.
         *             HSSF does not re-read the sheet or excepts.  Upon re-reading the sheet the value
         *             is incorrect or has not been replaced. <P>
         *
         */
        [Test]
        public void TestModifySimpleWithStyling()
        {
            HSSFWorkbook workbook = OpenSample("SimpleWithStyling.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);

            for (int k = 0; k < 4; k++)
            {
                ICell cell = sheet.GetRow(k).GetCell(0);

                cell.SetCellValue(new HSSFRichTextString(REPLACED));
            }


            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = workbook.GetSheetAt(0);
            for (int k = 0; k < 4; k++)
            {
                ICell cell = sheet.GetRow(k).GetCell(0);

                Assert.AreEqual(REPLACED, cell.RichStringCellValue.String);
            }
        }

        /**
         * TEST NAME:  Test Modify Employee Sheet<P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet with string values and replace
         *             them with other string values despite any styling.  In this release of HSSF styling will
         *             probably be lost and is NOT Tested.<P>
         * SUCCESS:    HSSF reads a sheet.  HSSF replaces the cell values with other cell values. HSSF
         *             Writes the sheet out to another file.  HSSF reads the result and ensures the value
         *             has been properly replaced.    <P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF does not Write the sheet or excepts.
         *             HSSF does not re-read the sheet or excepts.  Upon re-reading the sheet the value
         *             is incorrect or has not been replaced. <P>
         *
         */
        [Test]
        public void TestModifyEmployee()
        {
            HSSFWorkbook workbook = OpenSample("Employee.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            ICell cell = sheet.GetRow(3).GetCell(2);

            cell.SetCellValue(new HSSFRichTextString(LAST_NAME_VALUE));
            cell = sheet.GetRow(4).GetCell(2);
            cell.SetCellValue(new HSSFRichTextString(FIRST_NAME_VALUE));
            cell = sheet.GetRow(5).GetCell(2);
            cell.SetCellValue(new HSSFRichTextString(SSN_VALUE));

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = workbook.GetSheetAt(0);
            Assert.AreEqual(EMPLOYEE_INFORMATION, sheet.GetRow(1).GetCell(1).RichStringCellValue.String);
            Assert.AreEqual(LAST_NAME_VALUE, sheet.GetRow(3).GetCell(2).RichStringCellValue.String);
            Assert.AreEqual(FIRST_NAME_VALUE, sheet.GetRow(4).GetCell(2).RichStringCellValue.String);
            Assert.AreEqual(SSN_VALUE, sheet.GetRow(5).GetCell(2).RichStringCellValue.String);
        }

        /**
         * TEST NAME:  Test Read Sheet with an RK number<P>
         * OBJECTIVE:  Test that HSSF can read a simple spreadsheet with and RKRecord and correctly
         *             identify the cell as numeric and convert it to a NumberRecord.  <P>
         * SUCCESS:    HSSF reads a sheet.  HSSF returns that the cell is a numeric type cell.    <P>
         * FAILURE:    HSSF does not read a sheet or excepts.  HSSF incorrectly indentifies the cell<P>
         *
         */
        [Test]
        public void TestReadSheetWithRK()
        {
            HSSFWorkbook h = OpenSample("rk.xls");
            NPOI.SS.UserModel.ISheet s = h.GetSheetAt(0);
            ICell c = s.GetRow(0).GetCell(0);
            NPOI.SS.UserModel.CellType a = c.CellType;

            Assert.AreEqual(a, NPOI.SS.UserModel.CellType.Numeric);
        }

        /**
         * TEST NAME:  Test Write/Modify Sheet Simple <P>
         * OBJECTIVE:  Test that HSSF can Create a simple spreadsheet with numeric and string values,
         *             Remove some rows, yet still have a valid file/data.<P>
         * SUCCESS:    HSSF Creates a sheet.  Filesize Matches a known good.  NPOI.SS.UserModel.Sheet objects
         *             Last row, first row is Tested against the correct values (74,25).<P>
         * FAILURE:    HSSF does not Create a sheet or excepts.  Filesize does not Match the known good.
         *             NPOI.SS.UserModel.Sheet last row or first row is incorrect.             <P>
         *
         */
        [Test]
        public void TestWriteModifySheetMerged()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();

            for (int rownum = 0; rownum < 100; rownum++)
            {
                IRow r = s.CreateRow(rownum);

                for (int cellnum = 0; cellnum < 50; cellnum += 2)
                {
                    ICell c = r.CreateCell(cellnum);
                    c.SetCellValue(rownum * 10000 + cellnum
                                   + (((double)rownum / 1000)
                                      + ((double)cellnum / 10000)));
                    c = r.CreateCell(cellnum + 1);
                    c.SetCellValue(new HSSFRichTextString("TEST"));
                }
            }
            s.AddMergedRegion(new CellRangeAddress(0, 10, 0, 10));
            s.AddMergedRegion(new CellRangeAddress(30, 40, 5, 15));
            sanityChecker.CheckHSSFWorkbook(wb);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            s = wb.GetSheetAt(0);
            CellRangeAddress r1 = s.GetMergedRegion(0);
            CellRangeAddress r2 = s.GetMergedRegion(1);

            ConfirmRegion(new CellRangeAddress(0, 10, 0, 10), r1);
            ConfirmRegion(new CellRangeAddress(30, 40, 5, 15), r2);
        }

        private static void ConfirmRegion(CellRangeAddress ra, CellRangeAddress rb)
        {
            Assert.AreEqual(ra.FirstRow, rb.FirstRow);
            Assert.AreEqual(ra.LastRow, rb.LastRow);
            Assert.AreEqual(ra.FirstColumn, rb.FirstColumn);
            Assert.AreEqual(ra.LastColumn, rb.LastColumn);
        }

        /**
         * Test the backup field gets set as expected.
         */
        [Test]
        public void TestBackupRecord()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet();
            InternalWorkbook workbook = wb.Workbook;
            BackupRecord record = workbook.BackupRecord;

            Assert.AreEqual(0, record.Backup);
            Assert.IsFalse(wb.BackupFlag);
            wb.BackupFlag = (true);
            Assert.AreEqual(1, record.Backup);
            Assert.IsTrue(wb.BackupFlag);
            wb.BackupFlag = (false);
            Assert.AreEqual(0, record.Backup);
            Assert.IsFalse(wb.BackupFlag);
        }

        private class RecordCounter : RecordVisitor
        {
            private int _count;

            public RecordCounter()
            {
                _count = 0;
            }
            public int GetCount()
            {
                return _count;
            }
            public void VisitRecord(Record r)
            {
                if (r is LabelSSTRecord)
                {
                    _count++;
                }
            }
        }

        /**
         * This Tests is for bug [ #506658 ] Repeating output.
         *
         * We need to make sure only one LabelSSTRecord is produced.
         */
        [Test]
        public void TestRepeatingBug()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("Design Variants");
            IRow row = sheet.CreateRow(2);
            ICell cell = row.CreateCell(1);

            cell.SetCellValue(new HSSFRichTextString("Class"));
            cell = row.CreateCell(2);

            RecordCounter rc = new RecordCounter();
            sheet.Sheet.VisitContainedRecords(rc, 0);
            Assert.AreEqual(1, rc.GetCount());
        }

        [Test]
        public void TestManyRows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet();
            IRow row;
            ICell cell;
            int i, j;
            for (i = 0, j = 32771; j > 0; i++, j--)
            {
                row = sheet.CreateRow(i);
                cell = row.CreateCell(0);
                cell.SetCellValue(i);
            }
            sanityChecker.CheckHSSFWorkbook(workbook);
            Assert.AreEqual(32770, sheet.LastRowNum, "LAST ROW == 32770");
            cell = sheet.GetRow(32770).GetCell(0);
            double lastVal = cell.NumericCellValue;

            HSSFWorkbook wb = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            NPOI.SS.UserModel.ISheet s = wb.GetSheetAt(0);
            row = s.GetRow(32770);
            cell = row.GetCell(0);
            Assert.AreEqual(lastVal, cell.NumericCellValue, 0, "Value from last row == 32770");
            Assert.AreEqual(32770, s.LastRowNum, "LAST ROW == 32770");
        }

        /**
         * Generate a file to visually/programmatically verify repeating rows and cols made it
         */
        [Test]
        public void TestRepeatingColsRows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("Test Print Titles");

            IRow row = sheet.CreateRow(0);

            ICell cell = row.CreateCell(1);
            cell.SetCellValue(new HSSFRichTextString("hi"));


            workbook.SetRepeatingRowsAndColumns(0, 0, 1, 0, 0);

            string filepath = TempFile.GetTempFilePath("TestPrintTitles", ".xls");

            FileStream fileOut = new FileStream(filepath, FileMode.OpenOrCreate);
            workbook.Write(fileOut);
            fileOut.Close();

            Assert.IsTrue(File.Exists(filepath), "file exists");
        }

        [Test]
        public void TestNPOIBug6341()
        {
            {
                HSSFWorkbook workbook = OpenSample("Simple.xls");
                int i = workbook.ActiveSheetIndex;
            }
            {
                HSSFWorkbook workbook = OpenSample("blankworkbook.xls");
                int i = workbook.ActiveSheetIndex;
            }
            HSSFWorkbook workbook2 = OpenSample("blankworkbook.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook2.GetSheetAt(1);
        }
        [Test]
        public void TestRepeatingColsRowsMinusOne()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("Test Print Titles");

            IRow row = sheet.CreateRow(0);

            ICell cell = row.CreateCell(1);
            cell.SetCellValue(new HSSFRichTextString("hi"));


            workbook.SetRepeatingRowsAndColumns(0, -1, 1, -1, 0);

            FileInfo file = TempFile.CreateTempFile("testPrintTitlesA", ".xls");

            FileStream fileOut = new FileStream(file.FullName, FileMode.Create, FileAccess.ReadWrite);
            workbook.Write(fileOut);
            fileOut.Close();
            Assert.IsTrue(file.Exists, "file exists");
        }

        [Test]
        public void testAddMergedRegionWithRegion()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();

            for (int rownum = 0; rownum < 100; rownum++)
            {
                IRow r = s.CreateRow(rownum);

                for (int cellnum = 0; cellnum < 50; cellnum += 2)
                {
                    ICell c = r.CreateCell(cellnum);
                    c.SetCellValue(rownum * 10000 + cellnum
                                   + (((double)rownum / 1000)
                                      + ((double)cellnum / 10000)));
                    c = r.CreateCell(cellnum + 1);
                    c.SetCellValue(new HSSFRichTextString("TEST"));
                }
            }
            s.AddMergedRegion(new CellRangeAddress(0, (short)0, 10, (short)10));
            s.AddMergedRegion(new CellRangeAddress(30, 40, (short)5, (short)15));
            sanityChecker.CheckHSSFWorkbook(wb);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            s = wb.GetSheetAt(0);
            CellRangeAddress r1 = s.GetMergedRegion(0);
            CellRangeAddress r2 = s.GetMergedRegion(1);

            ConfirmRegion(new CellRangeAddress(0, 0, 10, 10), r1);
            ConfirmRegion(new CellRangeAddress(30, 40, 5, 15), r2);
        }
    }
}
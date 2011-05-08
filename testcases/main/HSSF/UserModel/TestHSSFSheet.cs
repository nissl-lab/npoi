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
    using System.IO;
    using System;
    using System.Configuration;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.SS.Util;
    using NPOI.DDF;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using NPOI.HSSF.Record.Aggregates;
    using TestCases.SS;
    using TestCases.SS.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    /**
     * Tests NPOI.SS.UserModel.Sheet.  This Test case is very incomplete at the moment.
     *
     *
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Andrew C. Oliver (acoliver apache org)
     */
    [TestClass]
    public class TestHSSFSheet : BaseTestSheet
    {
        public TestHSSFSheet():base(HSSFITestDataProvider.Instance)
        {
            
        }

        /**
         * Test the gridset field gets set as expected.
         */
        [TestMethod]
        public void TestBackupRecord()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;

            Assert.AreEqual(true, sheet.GridsetRecord.Gridset);
            s.IsGridsPrinted = true;
            Assert.AreEqual(false, sheet.GridsetRecord.Gridset);
        }

        /**
         * Test vertically centered output.
         */
        [TestMethod]
        public void TestVerticallyCenter()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;
            VCenterRecord record = sheet.PageSettings.VCenter;

            Assert.AreEqual(false, record.VCenter);
            s.VerticallyCenter = (true);
            Assert.AreEqual(true, record.VCenter);

            // wb.Write(new FileOutputStream("c:\\Test.xls"));
        }

        /**
         * Test horizontally centered output.
         */
        [TestMethod]
        public void TestHorizontallyCenter()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;
            HCenterRecord record = sheet.PageSettings.HCenter;

            Assert.AreEqual(false, record.HCenter);
            s.HorizontallyCenter = (true);
            Assert.AreEqual(true, record.HCenter);
        }


        /**
         * Test WSBboolRecord fields get set in the user model.
         */
        [TestMethod]
        public void TestWSBool()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;
            WSBoolRecord record =
                    (WSBoolRecord)sheet.FindFirstRecordBySid(WSBoolRecord.sid);

            // Check defaults
            Assert.AreEqual(true, record.AlternateExpression);
            Assert.AreEqual(true, record.AlternateFormula);
            Assert.AreEqual(false, record.Autobreaks);
            Assert.AreEqual(false, record.Dialog);
            Assert.AreEqual(false, record.DisplayGuts);
            Assert.AreEqual(true, record.FitToPage);
            Assert.AreEqual(false, record.RowSumsBelow);
            Assert.AreEqual(false, record.RowSumsRight);

            // Alter
            s.AlternativeExpression = (false);
            s.AlternativeFormula = (false);
            s.Autobreaks = true;
            s.Dialog = true;
            s.DisplayGuts = true;
            s.FitToPage = false;
            s.RowSumsBelow = true;
            s.RowSumsRight = true;

            // Check
            Assert.AreEqual(false, record.AlternateExpression);
            Assert.AreEqual(false, record.AlternateFormula);
            Assert.AreEqual(true, record.Autobreaks);
            Assert.AreEqual(true, record.Dialog);
            Assert.AreEqual(true, record.DisplayGuts);
            Assert.AreEqual(false, record.FitToPage);
            Assert.AreEqual(true, record.RowSumsBelow);
            Assert.AreEqual(true, record.RowSumsRight);
            Assert.AreEqual(false, s.AlternativeExpression);
            Assert.AreEqual(false, s.AlternativeFormula);
            Assert.AreEqual(true, s.Autobreaks);
            Assert.AreEqual(true, s.Dialog);
            Assert.AreEqual(true, s.DisplayGuts);
            Assert.AreEqual(false, s.FitToPage);
            Assert.AreEqual(true, s.RowSumsBelow);
            Assert.AreEqual(true, s.RowSumsRight);
        }
        [TestMethod]
        public void TestReadBooleans()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet("Test boolean");
            Row row = sheet.CreateRow(2);
            Cell cell = row.CreateCell(9);
            cell.SetCellValue(true);
            cell = row.CreateCell(11);
            cell.SetCellValue(true);

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);

            sheet = workbook.GetSheetAt(0);
            row = sheet.GetRow(2);
            Assert.IsNotNull(row);
            Assert.AreEqual(2, row.PhysicalNumberOfCells);
        }
        [TestMethod]
        public void TestRemoveZeroRow()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet("Sheet1");
            Row row = sheet.CreateRow(0);
            try
            {
                sheet.RemoveRow(row);
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Invalid row number (-1) outside allowable range (0..65535)"))
                {
                    throw new AssertFailedException("Identified bug 45367");
                }
                throw e;
            }
        }


        /**
         * Setting landscape and portrait stuff on existing sheets
         */
        [TestMethod]
        public void TestPrintSetupLandscapeExisting()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPageBreaks.xls");

            Assert.AreEqual(3, workbook.NumberOfSheets);

            NPOI.SS.UserModel.Sheet sheetL = workbook.GetSheetAt(0);
            NPOI.SS.UserModel.Sheet sheetPM = workbook.GetSheetAt(1);
            NPOI.SS.UserModel.Sheet sheetLS = workbook.GetSheetAt(2);

            // Check two aspects of the print setup
            Assert.IsFalse(sheetL.PrintSetup.Landscape);
            Assert.IsTrue(sheetPM.PrintSetup.Landscape);
            Assert.IsTrue(sheetLS.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(1, sheetPM.PrintSetup.Copies);
            Assert.AreEqual(1, sheetLS.PrintSetup.Copies);

            // Change one on each
            sheetL.PrintSetup.Landscape = (true);
            sheetPM.PrintSetup.Landscape = (false);
            sheetPM.PrintSetup.Copies = ((short)3);

            // Check taken
            Assert.IsTrue(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetPM.PrintSetup.Landscape);
            Assert.IsTrue(sheetLS.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(3, sheetPM.PrintSetup.Copies);
            Assert.AreEqual(1, sheetLS.PrintSetup.Copies);

            // Save and re-load, and Check still there
            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);

            Assert.IsTrue(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetPM.PrintSetup.Landscape);
            Assert.IsTrue(sheetLS.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(3, sheetPM.PrintSetup.Copies);
            Assert.AreEqual(1, sheetLS.PrintSetup.Copies);
        }
        [TestMethod]
        public void TestGroupRows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = workbook.CreateSheet();
            HSSFRow r1 = (HSSFRow)s.CreateRow(0);
            HSSFRow r2 = (HSSFRow)s.CreateRow(1);
            HSSFRow r3 = (HSSFRow)s.CreateRow(2);
            HSSFRow r4 = (HSSFRow)s.CreateRow(3);
            HSSFRow r5 = (HSSFRow)s.CreateRow(4);

            Assert.AreEqual(0, r1.OutlineLevel);
            Assert.AreEqual(0, r2.OutlineLevel);
            Assert.AreEqual(0, r3.OutlineLevel);
            Assert.AreEqual(0, r4.OutlineLevel);
            Assert.AreEqual(0, r5.OutlineLevel);

            s.GroupRow(2, 3);

            Assert.AreEqual(0, r1.OutlineLevel);
            Assert.AreEqual(0, r2.OutlineLevel);
            Assert.AreEqual(1, r3.OutlineLevel);
            Assert.AreEqual(1, r4.OutlineLevel);
            Assert.AreEqual(0, r5.OutlineLevel);

            // Save and re-Open
            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);

            s = workbook.GetSheetAt(0);
            r1 = (HSSFRow)s.GetRow(0);
            r2 = (HSSFRow)s.GetRow(1);
            r3 = (HSSFRow)s.GetRow(2);
            r4 = (HSSFRow)s.GetRow(3);
            r5 = (HSSFRow)s.GetRow(4);

            Assert.AreEqual(0, r1.OutlineLevel);
            Assert.AreEqual(0, r2.OutlineLevel);
            Assert.AreEqual(1, r3.OutlineLevel);
            Assert.AreEqual(1, r4.OutlineLevel);
            Assert.AreEqual(0, r5.OutlineLevel);
        }
        [TestMethod]
        public void TestGroupRowsExisting()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("NoGutsRecords.xls");

            NPOI.SS.UserModel.Sheet s = workbook.GetSheetAt(0);
            HSSFRow r1 = (HSSFRow)s.GetRow(0);
            HSSFRow r2 = (HSSFRow)s.GetRow(1);
            HSSFRow r3 = (HSSFRow)s.GetRow(2);
            HSSFRow r4 = (HSSFRow)s.GetRow(3);
            HSSFRow r5 = (HSSFRow)s.GetRow(4);
            HSSFRow r6 = (HSSFRow)s.GetRow(5);

            Assert.AreEqual(0, r1.OutlineLevel);
            Assert.AreEqual(0, r2.OutlineLevel);
            Assert.AreEqual(0, r3.OutlineLevel);
            Assert.AreEqual(0, r4.OutlineLevel);
            Assert.AreEqual(0, r5.OutlineLevel);
            Assert.AreEqual(0, r6.OutlineLevel);

            // This used to complain about lacking guts records
            s.GroupRow(2, 4);

            Assert.AreEqual(0, r1.OutlineLevel);
            Assert.AreEqual(0, r2.OutlineLevel);
            Assert.AreEqual(1, r3.OutlineLevel);
            Assert.AreEqual(1, r4.OutlineLevel);
            Assert.AreEqual(1, r5.OutlineLevel);
            Assert.AreEqual(0, r6.OutlineLevel);

            // Save and re-Open
            try
            {
                workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            }
            catch (OutOfMemoryException)
            {
                throw new AssertFailedException("Identified bug 39903");
            }

            s = workbook.GetSheetAt(0);
            r1 = (HSSFRow)s.GetRow(0);
            r2 = (HSSFRow)s.GetRow(1);
            r3 = (HSSFRow)s.GetRow(2);
            r4 = (HSSFRow)s.GetRow(3);
            r5 = (HSSFRow)s.GetRow(4);
            r6 = (HSSFRow)s.GetRow(5);

            Assert.AreEqual(0, r1.OutlineLevel);
            Assert.AreEqual(0, r2.OutlineLevel);
            Assert.AreEqual(1, r3.OutlineLevel);
            Assert.AreEqual(1, r4.OutlineLevel);
            Assert.AreEqual(1, r5.OutlineLevel);
            Assert.AreEqual(0, r6.OutlineLevel);
        }
        [TestMethod]
        public void TestGetDrawings()
        {
            HSSFWorkbook wb1c = HSSFTestDataSamples.OpenSampleWorkbook("WithChart.xls");
            HSSFWorkbook wb2c = HSSFTestDataSamples.OpenSampleWorkbook("WithTwoCharts.xls");

            // 1 chart sheet -> data on 1st, chart on 2nd
            Assert.IsNotNull(((HSSFSheet)wb1c.GetSheetAt(0)).DrawingPatriarch);
            Assert.IsNotNull(((HSSFSheet)wb1c.GetSheetAt(1)).DrawingPatriarch);
            Assert.IsFalse(((HSSFSheet)wb1c.GetSheetAt(0)).DrawingPatriarch.ContainsChart());
            Assert.IsTrue(((HSSFSheet)wb1c.GetSheetAt(1)).DrawingPatriarch.ContainsChart());

            // 2 chart sheet -> data on 1st, chart on 2nd+3rd
            Assert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(0)).DrawingPatriarch);
            Assert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(1)).DrawingPatriarch);
            Assert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(2)).DrawingPatriarch);
            Assert.IsFalse(((HSSFSheet)wb2c.GetSheetAt(0)).DrawingPatriarch.ContainsChart());
            Assert.IsTrue(((HSSFSheet)wb2c.GetSheetAt(1)).DrawingPatriarch.ContainsChart());
            Assert.IsTrue(((HSSFSheet)wb2c.GetSheetAt(2)).DrawingPatriarch.ContainsChart());
        }

        /**
 * Test that the ProtectRecord is included when creating or cloning a sheet
 */
        [TestMethod]
        public void TestCloneWithProtect()
        {
            String passwordA = "secrect";
            int expectedHashA = -6810;
            String passwordB = "admin";
            int expectedHashB = -14556;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet hssfSheet = (HSSFSheet)workbook.CreateSheet();
            hssfSheet.ProtectSheet(passwordA);

            Assert.AreEqual(expectedHashA, hssfSheet.Password);

            // Clone the sheet, and make sure the password hash is preserved
            HSSFSheet sheet2 = (HSSFSheet)workbook.CloneSheet(0);
            Assert.AreEqual(expectedHashA, sheet2.Password);

            // change the password on the first sheet
            hssfSheet.ProtectSheet(passwordB);
            Assert.AreEqual(expectedHashB, hssfSheet.Password);
            // but the cloned sheet's password should remain unchanged
            Assert.AreEqual(expectedHashA, sheet2.Password);
        }
        [TestMethod]
        public void TestProtectSheet()
        {
            short expected = unchecked((short)0xfef1);
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            s.ProtectSheet("abcdefghij");
            //WorksheetProtectionBlock pb = s.Sheet.ProtectionBlock;
            //Assert.IsTrue(pb.IsSheetProtected, "Protection should be on");
            //Assert.IsTrue(pb.IsObjectProtected, "object Protection should be on");
            //Assert.IsTrue(pb.IsScenarioProtected, "scenario Protection should be on");
            //Assert.AreEqual(expected, pb.PasswordHash, "well known value for top secret hash should be " + StringUtil.ToHexString(expected).Substring(4));
            Assert.IsTrue(s.Protect, "Protection should be on");
            Assert.IsTrue(s.ObjectProtect, "object Protection should be on");
            Assert.IsTrue(s.ScenarioProtect, "scenario Protection should be on");
            Assert.AreEqual(expected,s.Password, "well known value for top secret hash should be " + StringUtil.ToHexString(expected).Substring(4));
            
        }
        [TestMethod]
        public void TestProtectSheetRecordOrder_bug47363a()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            s.ProtectSheet("secret");
            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rc = new TestCases.HSSF.UserModel.RecordInspector.RecordCollector();
            s.Sheet.VisitContainedRecords(rc, 0);
            Record[] recs = rc.Records;
            int nRecs = recs.Length;
            if (recs[nRecs - 2] is PasswordRecord && recs[nRecs - 5] is DimensionsRecord)
            {
                throw new AssertFailedException("Identified bug 47363a - PASSWORD after DIMENSION");
            }
            // Check that protection block is together, and before DIMENSION
            ConfirmRecordClass(recs, nRecs - 5, typeof(DimensionsRecord));
            ConfirmRecordClass(recs, nRecs - 10, typeof(ProtectRecord));
            ConfirmRecordClass(recs, nRecs - 9, typeof(ObjectProtectRecord));
            ConfirmRecordClass(recs, nRecs - 8, typeof(ScenarioProtectRecord));
            ConfirmRecordClass(recs, nRecs - 7, typeof(PasswordRecord));
        }
        private void ConfirmRecordClass(Record[] recs, int index, Type cls)
        {
            if (recs.Length <= index)
            {
                throw new AssertFailedException("Expected (" + cls.Name + ") at index "
                        + index + " but array length is " + recs.Length + ".");
            }
            Assert.AreEqual(cls, recs[index].GetType());
        }
        [TestMethod]
        public void TestZoom()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet();
            Assert.AreEqual(-1, sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid));
            sheet.SetZoom(3, 4);
            Assert.IsTrue(sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid) > 0);
            SCLRecord sclRecord = (SCLRecord)sheet.Sheet.FindFirstRecordBySid(SCLRecord.sid);
            Assert.AreEqual(3, sclRecord.Numerator);
            Assert.AreEqual(4, sclRecord.Denominator);

            int sclLoc = sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid);
            int window2Loc = sheet.Sheet.FindFirstRecordLocBySid(WindowTwoRecord.sid);
            Assert.IsTrue(sclLoc == window2Loc + 1);
        }



        /**
         * Make sure the excel file loads work
         *
         */
        [TestMethod]
        public void TestPageBreakFiles()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPageBreaks.xls");

            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);
            Assert.IsNotNull(sheet);

            Assert.AreEqual(1, sheet.RowBreaks.Length, "1 row page break");
            Assert.AreEqual(1, sheet.ColumnBreaks.Length, "1 column page break");

            Assert.IsTrue(sheet.IsRowBroken(22), "No row page break");
            Assert.IsTrue(sheet.IsColumnBroken((short)4), "No column page break");

            sheet.SetRowBreak(10);
            sheet.SetColumnBreak((short)13);

            Assert.AreEqual(2, sheet.RowBreaks.Length, "row breaks number");
            Assert.AreEqual(2, sheet.ColumnBreaks.Length, "column breaks number");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);

            Assert.IsTrue(sheet.IsRowBroken(22), "No row page break");
            Assert.IsTrue(sheet.IsColumnBroken((short)4), "No column page break");

            Assert.AreEqual(2, sheet.RowBreaks.Length, "row breaks number");
            Assert.AreEqual(2, sheet.ColumnBreaks.Length, "column breaks number");
        }
        [TestMethod]
        public void TestDBCSName()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("DBCSSheetName.xls");
            wb.GetSheetAt(1);
            Assert.AreEqual(wb.GetSheetName(1), "\u090f\u0915", "DBCS Sheet Name 2");
            Assert.AreEqual(wb.GetSheetName(0), "\u091c\u093e", "DBCS Sheet Name 1");
        }

        /**
         * Testing newly Added method that exposes the WINDOW2.toprow
         * parameter to allow setting the toprow in the visible view
         * of the sheet when it is first Opened.
         */
        [TestMethod]
        public void TestTopRow()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPageBreaks.xls");

            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);
            Assert.IsNotNull(sheet);

            short toprow = (short)100;
            short leftcol = (short)50;
            sheet.ShowInPane(toprow, leftcol);
            Assert.AreEqual(toprow, sheet.TopRow, "NPOI.SS.UserModel.Sheet.GetTopRow()");
            Assert.AreEqual(leftcol, sheet.LeftCol, "NPOI.SS.UserModel.Sheet.GetLeftCol()");
        }

        /** cell with formula becomes null on cloning a sheet*/
        [TestMethod]
        public void Test35084()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet("Sheet1");
            Row r = s.CreateRow(0);
            r.CreateCell(0).SetCellValue(1);
            r.CreateCell(1).CellFormula = ("A1*2");
            NPOI.SS.UserModel.Sheet s1 = wb.CloneSheet(0);
            r = s1.GetRow(0);
            Assert.AreEqual(r.GetCell(0).NumericCellValue, 1, 1); // sanity Check
            Assert.IsNotNull(r.GetCell(1));
            Assert.AreEqual(r.GetCell(1).CellFormula, "A1*2");
        }

        /** Test that new default column styles get applied */
        [TestMethod]
        public void TestDefaultColumnStyle()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.CellStyle style = wb.CreateCellStyle();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            s.SetDefaultColumnStyle((short)0, style);
            Row r = s.CreateRow(0);
            Cell c = r.CreateCell(0);
            Assert.AreEqual(style.Index, c.CellStyle.Index, "style should Match");
        }


        /**
         *
         */
        [TestMethod]
        public void TestAddEmptyRow()
        {
            //try to Add 5 empty rows to a new sheet
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet = workbook.CreateSheet();
            for (int i = 0; i < 5; i++)
            {
                sheet.CreateRow(i);
            }

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);

            //try Adding empty rows in an existing worksheet
            workbook = HSSFTestDataSamples.OpenSampleWorkbook("Simple.xls");

            sheet = workbook.GetSheetAt(0);
            for (int i = 3; i < 10; i++) sheet.CreateRow(i);

            workbook = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
        }
        [TestMethod]
        public void TestAutoSizeColumn()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("43902.xls");
            String sheetName = "my sheet";
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet();

            // Can't use literal numbers for column sizes, as
            //  will come out with different values on different
            //  machines based on the fonts available.
            // So, we use ranges, which are pretty large, but
            //  thankfully don't overlap!
            int minWithRow1And2 = 6400;
            int maxWithRow1And2 = 7800;
            int minWithRow1Only = 2750;
            int maxWithRow1Only = 3300;

            // autoSize the first column and check its size before the merged region (1,0,1,1) is set:
            // it has to be based on the 2nd row width
            sheet.AutoSizeColumn(0);
            Assert.IsTrue(sheet.GetColumnWidth(0) >= minWithRow1And2, "Column autosized with only one row: wrong width");
            Assert.IsTrue(sheet.GetColumnWidth(0) <= maxWithRow1And2, "Column autosized with only one row: wrong width");

            //Create a region over the 2nd row and auto size the first column
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            sheet.AutoSizeColumn(0);
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            // Check that the autoSized column width has ignored the 2nd row
            // because it is included in a merged region (Excel like behavior)
            NPOI.SS.UserModel.Sheet sheet2 = wb2.GetSheet(sheetName);
            Assert.IsTrue(sheet2.GetColumnWidth(0) >= minWithRow1Only);
            Assert.IsTrue(sheet2.GetColumnWidth(0) <= maxWithRow1Only);

            // Remove the 2nd row merged region and Check that the 2nd row value is used to the AutoSizeColumn width
            sheet2.RemoveMergedRegion(1);
            sheet2.AutoSizeColumn(0);
            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            NPOI.SS.UserModel.Sheet sheet3 = wb3.GetSheet(sheetName);
            Assert.IsTrue(sheet3.GetColumnWidth(0) >= minWithRow1And2);
            Assert.IsTrue(sheet3.GetColumnWidth(0) <= maxWithRow1And2);
        }

        ///**
        // * Setting ForceFormulaRecalculation on sheets
        // */
        //public void TestForceRecalculation()
        //{
        //    HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("UncalcedRecord.xls");

        //    NPOI.SS.UserModel.Sheet sheet = workbook.GetSheetAt(0);
        //    NPOI.SS.UserModel.Sheet sheet2 = workbook.GetSheetAt(0);
        //    Row row = sheet.GetRow(0);
        //    row.CreateCell(0).SetCellValue(5);
        //    row.CreateCell(1).SetCellValue(8);
        //    Assert.IsFalse(sheet.ForceFormulaRecalculation);
        //    Assert.IsFalse(sheet2.ForceFormulaRecalculation);

        //    string path1 = ConfigurationSettings.AppSettings["java.io.tmpdir"] + "/uncalced_err.xls";
        //    MemoryStream ms=new MemoryStream(File.ReadAllBytes(path1));
        //    File.Delete(path1);

        //    FileStream fout = File.Open(path1, FileMode.OpenOrCreate);
        //    workbook.Write(fout);
        //    fout.Close();
        //    sheet.ForceFormulaRecalculation = (true);
        //    Assert.IsTrue(sheet.ForceFormulaRecalculation);
        //    // Save and manually verify that on column C we have 0, value in template

        //    string path2 = ConfigurationSettings.AppSettings["java.io.tmpdir"] + "/uncalced_succ.xls";

        //    // Save and manually verify that on column C we have now 13, calculated value
        //    tempFile = new File();
        //    tempFile.delete();
        //    fout = new FileOutputStream(tempFile);
        //    workbook.Write(fout);
        //    fout.Close();

        //    // Try it can be Opened
        //    HSSFWorkbook wb2 = new HSSFWorkbook(new FileInputStream(tempFile));

        //    // And Check correct sheet settings found
        //    sheet = wb2.GetSheetAt(0);
        //    sheet2 = wb2.GetSheetAt(1);
        //    Assert.IsTrue(sheet.ForceFormulaRecalculation);
        //    Assert.IsFalse(sheet2.ForceFormulaRecalculation);

        //    // Now turn if back off again
        //    sheet.ForceFormulaRecalculation=(false);

        //    fout = new FileOutputStream(tempFile);
        //    wb2.Write(fout);
        //    fout.Close();
        //    wb2 = new HSSFWorkbook(new FileInputStream(tempFile));

        //    Assert.IsFalse(wb2.GetSheetAt(0).ForceFormulaRecalculation);
        //    Assert.IsFalse(wb2.GetSheetAt(1).ForceFormulaRecalculation);
        //    Assert.IsFalse(wb2.GetSheetAt(2).ForceFormulaRecalculation);

        //    // Now Add a new sheet, and Check things work
        //    //  with old ones unset, new one set
        //    NPOI.SS.UserModel.Sheet s4 = wb2.CreateSheet();
        //    s4.ForceFormulaRecalculation=(true);

        //    Assert.IsFalse(sheet.ForceFormulaRecalculation);
        //    Assert.IsFalse(sheet2.ForceFormulaRecalculation);
        //    Assert.IsTrue(s4.ForceFormulaRecalculation);

        //    fout = new FileOutputStream(tempFile);
        //    wb2.Write(fout);
        //    fout.Close();

        //    HSSFWorkbook wb3 = new HSSFWorkbook(new FileInputStream(tempFile));
        //    Assert.IsFalse(wb3.GetSheetAt(0).ForceFormulaRecalculation);
        //    Assert.IsFalse(wb3.GetSheetAt(1).ForceFormulaRecalculation);
        //    Assert.IsFalse(wb3.GetSheetAt(2).ForceFormulaRecalculation);
        //    Assert.IsTrue(wb3.GetSheetAt(3).ForceFormulaRecalculation);
        //}
 

        /**
         * Some utilities Write Excel files without the ROW records.
         * Excel, ooo, and google docs are OK with this.
         * Now POI is too.
         */
        [TestMethod]
        public void TestMissingRowRecords_bug41187()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex41187-19267.xls");

            NPOI.SS.UserModel.Sheet sheet = wb.GetSheetAt(0);
            Row row = sheet.GetRow(0);
            if (row == null)
            {
                throw new AssertFailedException("Identified bug 41187 a");
            }
            if (row.Height == 0)
            {
                throw new AssertFailedException("Identified bug 41187 b");
            }
            Assert.AreEqual("Hi Excel!", row.GetCell(0).RichStringCellValue.String);
            // Check row height for 'default' flag
            Assert.AreEqual((short)0xFF, row.Height);

            HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }

        /**
         * If we Clone a sheet containing drawings,
         * we must allocate a new ID of the drawing Group and re-Create shape IDs
         *
         * See bug #45720.
         */
        [TestMethod]
        public void TestCloneSheetWithDrawings()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("45720.xls");

            HSSFSheet sheet1 = (HSSFSheet)wb1.CreateSheet();

            wb1.Workbook.FindDrawingGroup();
            DrawingManager2 dm1 = wb1.Workbook.DrawingManager;

            wb1.CloneSheet(0);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb2.Workbook.FindDrawingGroup();
            DrawingManager2 dm2 = wb2.Workbook.DrawingManager;

            //Check EscherDggRecord - a workbook-level registry of drawing objects
            Assert.AreEqual(dm1.GetDgg().MaxDrawingGroupId + 1, dm2.GetDgg().MaxDrawingGroupId);

            sheet1 = (HSSFSheet)wb2.GetSheetAt(0);
            HSSFSheet sheet2 = (HSSFSheet)wb2.GetSheetAt(1);

            //Check that id of the drawing Group was updated
            EscherDgRecord dg1 = (EscherDgRecord)sheet1.DrawingEscherAggregate.FindFirstWithId(EscherDgRecord.RECORD_ID);
            EscherDgRecord dg2 = (EscherDgRecord)sheet2.DrawingEscherAggregate.FindFirstWithId(EscherDgRecord.RECORD_ID);
            int dg_id_1 = dg1.Options >> 4;
            int dg_id_2 = dg2.Options >> 4;
            Assert.AreEqual(dg_id_1 + 1, dg_id_2);

            //TODO: Check shapeId in the Cloned sheet
        }

        /**
         * POI now (Sep 2008) allows sheet names longer than 31 chars (for other apps besides Excel).
         * Since Excel silently truncates to 31, make sure that POI enforces uniqueness on the first
         * 31 chars. 
         */
        [TestMethod]
        public void TestLongSheetNames()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            String SAME_PREFIX = "A123456789B123456789C123456789"; // 30 chars

            wb.CreateSheet(SAME_PREFIX + "Dxxxx");
            try
            {
                wb.CreateSheet(SAME_PREFIX + "Dyyyy"); // identical up to the 32nd char
                throw new AssertFailedException("Expected exception not thrown");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The workbook already contains a sheet of this name", e.Message);
            }
            wb.CreateSheet(SAME_PREFIX + "Exxxx"); // OK - differs in the 31st char
        }
        [TestMethod]
        public void TestArabic()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();

            Assert.AreEqual(false, s.IsArabic);
            s.IsArabic = true;
            Assert.AreEqual(true, s.IsArabic);
        }

    }
}
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

    using TestCases.HSSF;
    using NPOI.HSSF.Record.Aggregates;
    using TestCases.SS;
    using TestCases.SS.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.HSSF.Record.AutoFilter;
    using System.Collections.Generic;
    using System.Collections;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NUnit.Framework;
    using NPOI.SS;

    /**
     * Tests NPOI.SS.UserModel.Sheet.  This Test case is very incomplete at the moment.
     *
     *
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Andrew C. Oliver (acoliver apache org)
     */
    [TestFixture]
    public class TestHSSFSheet : BaseTestSheet
    {
        public TestHSSFSheet()
            : base(HSSFITestDataProvider.Instance)
        {

        }

        /**
     * Test for Bugzilla #29747.
     * Moved from TestHSSFWorkbook#testSetRepeatingRowsAndColumns().
     */
        [Test]
        public void TestSetRepeatingRowsAndColumnsBug29747()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet();
            wb.CreateSheet();
            HSSFSheet sheet2 = (HSSFSheet)wb.CreateSheet();
            sheet2.RepeatingRows = (CellRangeAddress.ValueOf("1:2"));
            NameRecord nameRecord = wb.Workbook.GetNameRecord(0);
            Assert.AreEqual(3, nameRecord.SheetNumber);

            wb.Close();
        }
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }
        [Test]
        public void TestTestGetSetMargin()
        {
            BaseTestGetSetMargin(new double[] { 0.75, 0.75, 1.0, 1.0, 0.3, 0.3 });
        }

        /**
         * Test the gridset field gets set as expected.
         */
        [Test]
        public void TestBackupRecord()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;

            Assert.IsTrue(sheet.GridsetRecord.Gridset);
#pragma warning disable 0618 //  warning CS0618: 'NPOI.HSSF.UserModel.HSSFSheet.IsGridsPrinted' is obsolete: 'Please use IsPrintGridlines instead'
            s.IsGridsPrinted = true; // <- this is marked obsolete, but using "s.IsPrintGridlines = true;" makes this test fail 8-(
#pragma warning restore 0618
            Assert.IsFalse(sheet.GridsetRecord.Gridset);

            wb.Close();
        }

        /**
         * Test vertically centered output.
         */
        [Test]
        public void TestVerticallyCenter()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;
            VCenterRecord record = sheet.PageSettings.VCenter;

            Assert.IsFalse(record.VCenter);
            Assert.IsFalse(s.VerticallyCenter);
            s.VerticallyCenter = (true);
            Assert.IsTrue(record.VCenter);
            Assert.IsTrue(s.VerticallyCenter);

            wb.Close();
        }

        /**
         * Test horizontally centered output.
         */
        [Test]
        public void TestHorizontallyCenter()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;
            HCenterRecord record = sheet.PageSettings.HCenter;

            Assert.IsFalse(record.HCenter);
            Assert.IsFalse(s.HorizontallyCenter);
            s.HorizontallyCenter = (true);
            Assert.IsTrue(record.HCenter);
            Assert.IsTrue(s.HorizontallyCenter);

            wb.Close();
        }


        /**
         * Test WSBboolRecord fields get set in the user model.
         */
        [Test]
        public void TestWSBool()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            InternalSheet sheet = s.Sheet;
            WSBoolRecord record =
                    (WSBoolRecord)sheet.FindFirstRecordBySid(WSBoolRecord.sid);

            // Check defaults
            Assert.IsFalse(record.AlternateExpression);
            Assert.IsFalse(record.AlternateFormula);
            Assert.IsFalse(record.Autobreaks);
            Assert.IsFalse(record.Dialog);
            Assert.IsFalse(record.DisplayGuts);
            Assert.IsTrue(record.FitToPage);
            Assert.IsFalse(record.RowSumsBelow);
            Assert.IsFalse(record.RowSumsRight);

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
            Assert.IsFalse(record.AlternateExpression);
            Assert.IsFalse(record.AlternateFormula);
            Assert.IsTrue(record.Autobreaks);
            Assert.IsTrue(record.Dialog);
            Assert.IsTrue(record.DisplayGuts);
            Assert.IsFalse(record.FitToPage);
            Assert.IsTrue(record.RowSumsBelow);
            Assert.IsTrue(record.RowSumsRight);
            Assert.IsFalse(s.AlternativeExpression);
            Assert.IsFalse(s.AlternativeFormula);
            Assert.IsTrue(s.Autobreaks);
            Assert.IsTrue(s.Dialog);
            Assert.IsTrue(s.DisplayGuts);
            Assert.IsFalse(s.FitToPage);
            Assert.IsTrue(s.RowSumsBelow);
            Assert.IsTrue(s.RowSumsRight);

            wb.Close();
        }
        [Test]
        [Ignore("not fount in poi")]
        public void TestReadBooleans()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb1.CreateSheet("Test boolean");
            IRow row = sheet.CreateRow(2);
            ICell cell = row.CreateCell(9);
            cell.SetCellValue(true);
            cell = row.CreateCell(11);
            cell.SetCellValue(true);

            IWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            row = sheet.GetRow(2);
            Assert.IsNotNull(row);
            Assert.AreEqual(2, row.PhysicalNumberOfCells);

            wb2.Close();
        }
        [Test]
        [Ignore("not fount in poi")]
        public void TestRemoveZeroRow()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            try
            {
                sheet.RemoveRow(row);
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Invalid row number (-1) outside allowable range (0..65535)"))
                {
                    throw new AssertionException("Identified bug 45367");
                }
                throw e;
            }
        }


        /**
         * Setting landscape and portrait stuff on existing sheets
         */
        [Test]
        public void TestPrintSetupLandscapeExisting()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPageBreaks.xls");

            Assert.AreEqual(3, wb1.NumberOfSheets);

            NPOI.SS.UserModel.ISheet sheetL = wb1.GetSheetAt(0);
            NPOI.SS.UserModel.ISheet sheetPM = wb1.GetSheetAt(1);
            NPOI.SS.UserModel.ISheet sheetLS = wb1.GetSheetAt(2);

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
            IWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheetL = wb1.GetSheetAt(0);
            sheetPM = wb1.GetSheetAt(1);
            sheetLS = wb1.GetSheetAt(2);

            Assert.IsTrue(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetPM.PrintSetup.Landscape);
            Assert.IsTrue(sheetLS.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(3, sheetPM.PrintSetup.Copies);
            Assert.AreEqual(1, sheetLS.PrintSetup.Copies);

            wb2.Close();
        }
        [Test]
        public void TestGroupRows()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb1.CreateSheet();
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
            IWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            s = wb2.GetSheetAt(0);
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

            wb2.Close();
        }
        [Test]
        public void TestGroupRowsExisting()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("NoGutsRecords.xls");

            NPOI.SS.UserModel.ISheet s = wb1.GetSheetAt(0);
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
            HSSFWorkbook wb2;
            try
            {
                wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            }
            catch (OutOfMemoryException)
            {
                throw new AssertionException("Identified bug 39903");
            }

            s = wb2.GetSheetAt(0);
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

            wb2.Close();
            wb1.Close();
        }
        [Test]
        public void TestCreateDrawings()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            HSSFPatriarch p1 = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            HSSFPatriarch p2 = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            Assert.AreSame(p1, p2);
            workbook.Close();
        }
        [Test]
        public void TestGetDrawings()
        {
            HSSFWorkbook wb1c = HSSFTestDataSamples.OpenSampleWorkbook("WithChart.xls");
            HSSFWorkbook wb2c = HSSFTestDataSamples.OpenSampleWorkbook("WithTwoCharts.xls");

            // 1 chart sheet -> data on 1st, chart on 2nd
            Assert.IsNotNull(((HSSFSheet)wb1c.GetSheetAt(0)).DrawingPatriarch);
            Assert.IsNotNull(((HSSFSheet)wb1c.GetSheetAt(1)).DrawingPatriarch);
            Assert.IsFalse((((HSSFSheet)wb1c.GetSheetAt(0)).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            Assert.IsTrue((((HSSFSheet)wb1c.GetSheetAt(1)).DrawingPatriarch as HSSFPatriarch).ContainsChart());

            // 2 chart sheet -> data on 1st, chart on 2nd+3rd
            Assert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(0)).DrawingPatriarch);
            Assert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(1)).DrawingPatriarch);
            Assert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(2)).DrawingPatriarch);
            Assert.IsFalse((((HSSFSheet)wb2c.GetSheetAt(0)).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            Assert.IsTrue((((HSSFSheet)wb2c.GetSheetAt(1)).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            Assert.IsTrue((((HSSFSheet)wb2c.GetSheetAt(2)).DrawingPatriarch as HSSFPatriarch).ContainsChart());

            wb2c.Close();
            wb1c.Close();
        }

        /**
 * Test that the ProtectRecord is included when creating or cloning a sheet
 */
        [Test]
        public void TestCloneWithProtect()
        {
            String passwordA = "secrect";
            int expectedHashA = -6810;
            String passwordB = "admin";
            int expectedHashB = -14556;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet hssfSheet = (HSSFSheet)workbook.CreateSheet();
            Assert.IsFalse(hssfSheet.ObjectProtect);
            hssfSheet.ProtectSheet(passwordA);
            Assert.IsTrue(hssfSheet.ObjectProtect);
            Assert.AreEqual(expectedHashA, hssfSheet.Password);
            Assert.AreEqual(expectedHashA, hssfSheet.Sheet.ProtectionBlock.PasswordHash);
            

            // Clone the sheet, and make sure the password hash is preserved
            HSSFSheet sheet2 = (HSSFSheet)workbook.CloneSheet(0);
            Assert.IsTrue(hssfSheet.ObjectProtect);
            Assert.AreEqual(expectedHashA, sheet2.Sheet.ProtectionBlock.PasswordHash);

            // change the password on the first sheet
            hssfSheet.ProtectSheet(passwordB);
            Assert.IsTrue(hssfSheet.ObjectProtect);
            Assert.AreEqual(expectedHashB, hssfSheet.Sheet.ProtectionBlock.PasswordHash);
            Assert.AreEqual(expectedHashB, hssfSheet.Password);
            // but the cloned sheet's password should remain unchanged
            Assert.AreEqual(expectedHashA, sheet2.Password);

            workbook.Close();
        }
        [Test]
        public new void TestProtectSheet()
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
            Assert.AreEqual(expected, s.Password, "well known value for top secret hash should be " + StringUtil.ToHexString(expected).Substring(4));

            wb.Close();

        }
        [Test]
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
                Assert.Fail("Identified bug 47363a - PASSWORD after DIMENSION");
            }
            // Check that protection block is together, and before DIMENSION
            //ConfirmRecordClass(recs, nRecs - 4, typeof(DimensionsRecord));
            //ConfirmRecordClass(recs, nRecs - 9, typeof(ProtectRecord));
            //ConfirmRecordClass(recs, nRecs - 8, typeof(ObjectProtectRecord));
            //ConfirmRecordClass(recs, nRecs - 7, typeof(ScenarioProtectRecord));
            //ConfirmRecordClass(recs, nRecs - 6, typeof(PasswordRecord));

            ConfirmRecordClass(recs, nRecs - 5, typeof(DimensionsRecord));
            ConfirmRecordClass(recs, nRecs - 10, typeof(ProtectRecord));
            ConfirmRecordClass(recs, nRecs - 9, typeof(ObjectProtectRecord));
            ConfirmRecordClass(recs, nRecs - 8, typeof(ScenarioProtectRecord));
            ConfirmRecordClass(recs, nRecs - 7, typeof(PasswordRecord));

            wb.Close();
        }
        private void ConfirmRecordClass(Record[] recs, int index, Type cls)
        {
            if (recs.Length <= index)
            {
                throw new AssertionException("Expected (" + cls.Name + ") at index "
                        + index + " but array length is " + recs.Length + ".");
            }
            Assert.AreEqual(cls, recs[index].GetType());
        }

        /**
    * There should be no problem with Adding data validations After sheet protection
    */
        [Test]
        public void TestDvProtectionOrder_bug47363b()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");
            sheet.ProtectSheet("secret");

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint dvc = dataValidationHelper.CreateintConstraint(OperatorType.BETWEEN, "10", "100");
            CellRangeAddressList numericCellAddressList = new CellRangeAddressList(0, 0, 1, 1);
            IDataValidation dv = dataValidationHelper.CreateValidation(dvc, numericCellAddressList);
            try
            {
                sheet.AddValidationData(dv);
            }
            catch (InvalidOperationException e)
            {
                String expMsg = "Unexpected (NPOI.HSSF.Record.PasswordRecord) while looking for DV Table insert pos";
                if (expMsg.Equals(e.Message))
                {
                    Assert.Fail("Identified bug 47363b");
                }
                workbook.Close();
                throw e;
            }
            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rc;
            rc = new RecordInspector.RecordCollector();
            ((HSSFSheet)sheet).Sheet.VisitContainedRecords(rc, 0);
            int nRecsWithProtection = rc.Records.Length;

            sheet.ProtectSheet(null);
            rc = new RecordInspector.RecordCollector();
            ((HSSFSheet)sheet).Sheet.VisitContainedRecords(rc, 0);
            int nRecsWithoutProtection = rc.Records.Length;

            Assert.AreEqual(4, nRecsWithProtection - nRecsWithoutProtection);

            workbook.Close();
        }
        [Test]
        public void TestZoom()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet();
            Assert.AreEqual(-1, sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid));
            sheet.SetZoom(75);
            Assert.IsTrue(sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid) > 0);
            SCLRecord sclRecord = (SCLRecord)sheet.Sheet.FindFirstRecordBySid(SCLRecord.sid);
            Assert.AreEqual(75, 100 * sclRecord.Numerator / sclRecord.Denominator);

            int sclLoc = sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid);
            int window2Loc = sheet.Sheet.FindFirstRecordLocBySid(WindowTwoRecord.sid);
            Assert.IsTrue(sclLoc == window2Loc + 1);

            // verify limits
            try
            {
                sheet.SetZoom(0);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Numerator must be greater than 0 and less than 65536", e.Message);
            }
            try
            {
                sheet.SetZoom(65536);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Numerator must be greater than 0 and less than 65536", e.Message);
            }
            try
            {
                sheet.SetZoom(2, 0);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Denominator must be greater than 0 and less than 65536", e.Message);
            }
            try
            {
                sheet.SetZoom(2, 65536);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Denominator must be greater than 0 and less than 65536", e.Message);
            }

            wb.Close();
        }



        /**
         * Make sure the excel file loads work
         *
         */
        [Test]
        public void TestPageBreakFiles()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPageBreaks.xls");

            NPOI.SS.UserModel.ISheet sheet = wb1.GetSheetAt(0);
            Assert.IsNotNull(sheet);

            Assert.AreEqual(1, sheet.RowBreaks.Length, "1 row page break");
            Assert.AreEqual(1, sheet.ColumnBreaks.Length, "1 column page break");

            Assert.IsTrue(sheet.IsRowBroken(22), "No row page break");
            Assert.IsTrue(sheet.IsColumnBroken((short)4), "No column page break");

            sheet.SetRowBreak(10);
            sheet.SetColumnBreak((short)13);

            Assert.AreEqual(2, sheet.RowBreaks.Length, "row breaks number");
            Assert.AreEqual(2, sheet.ColumnBreaks.Length, "column breaks number");

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);

            Assert.IsTrue(sheet.IsRowBroken(22), "No row page break");
            Assert.IsTrue(sheet.IsColumnBroken((short)4), "No column page break");

            Assert.AreEqual(2, sheet.RowBreaks.Length, "row breaks number");
            Assert.AreEqual(2, sheet.ColumnBreaks.Length, "column breaks number");

            wb2.Close();
        }
        [Test]
        public void TestDBCSName()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("DBCSSheetName.xls");
            wb.GetSheetAt(1);
            Assert.AreEqual(wb.GetSheetName(1), "\u090f\u0915", "DBCS Sheet Name 2");
            Assert.AreEqual(wb.GetSheetName(0), "\u091c\u093e", "DBCS Sheet Name 1");
            wb.Close();
        }

        /**
         * Testing newly Added method that exposes the WINDOW2.toprow
         * parameter to allow setting the toprow in the visible view
         * of the sheet when it is first Opened.
         */
        [Test]
        public void TestTopRow()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPageBreaks.xls");

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
            Assert.IsNotNull(sheet);

            short toprow = (short)100;
            short leftcol = (short)50;
            sheet.ShowInPane(toprow, leftcol);
            Assert.AreEqual(toprow, sheet.TopRow, "NPOI.SS.UserModel.Sheet.GetTopRow()");
            Assert.AreEqual(leftcol, sheet.LeftCol, "NPOI.SS.UserModel.Sheet.GetLeftCol()");

            wb.Close();
        }

        /** cell with formula becomes null on cloning a sheet*/
        [Test]
        public void Test35084()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet("Sheet1");
            IRow r = s.CreateRow(0);
            r.CreateCell(0).SetCellValue(1);
            r.CreateCell(1).CellFormula = ("A1*2");
            NPOI.SS.UserModel.ISheet s1 = wb.CloneSheet(0);
            r = s1.GetRow(0);
            Assert.AreEqual(r.GetCell(0).NumericCellValue, 1, 1); // sanity Check
            Assert.IsNotNull(r.GetCell(1));
            Assert.AreEqual(r.GetCell(1).CellFormula, "A1*2");

            wb.Close();
        }


        /**
         *
         */
        [Test]
        public void TestAddEmptyRow()
        {
            //try to Add 5 empty rows to a new sheet
            HSSFWorkbook wb1 = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb1.CreateSheet();
            for (int i = 0; i < 5; i++)
            {
                sheet.CreateRow(i);
            }

            HSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();
            wb1.Close();

            //try Adding empty rows in an existing worksheet
            HSSFWorkbook wb2 = HSSFTestDataSamples.OpenSampleWorkbook("Simple.xls");

            sheet = wb2.GetSheetAt(0);
            for (int i = 3; i < 10; i++) sheet.CreateRow(i);

            HSSFTestDataSamples.WriteOutAndReadBack(wb2).Close();

            wb2.Close();
        }
        [Test]
        public void TestAutoSizeColumn()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("43902.xls");
            String sheetName = "my sheet";
            HSSFSheet sheet = (HSSFSheet)wb1.GetSheet(sheetName);

            // Can't use literal numbers for column sizes, as
            //  will come out with different values on different
            //  machines based on the fonts available.
            // So, we use ranges, which are pretty large, but
            //  thankfully don't overlap!
            int minWithRow1And2 = 6400;
            int maxWithRow1And2 = 7800;
            int minWithRow1Only = 2730;
            int maxWithRow1Only = 3300;

            // autoSize the first column and check its size before the merged region (1,0,1,1) is set:
            // it has to be based on the 2nd row width
            sheet.AutoSizeColumn(0);
            Assert.IsTrue(sheet.GetColumnWidth(0) >= minWithRow1And2, "Column autosized with only one row: wrong width");
            Assert.IsTrue(sheet.GetColumnWidth(0) <= maxWithRow1And2, "Column autosized with only one row: wrong width");

            //Create a region over the 2nd row and auto size the first column
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            Assert.IsNotNull(sheet.GetMergedRegion(0));
            sheet.AutoSizeColumn(0);
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);

            // Check that the autoSized column width has ignored the 2nd row
            // because it is included in a merged region (Excel like behavior)
            NPOI.SS.UserModel.ISheet sheet2 = wb2.GetSheet(sheetName);
            Assert.IsTrue(sheet2.GetColumnWidth(0) >= minWithRow1Only, $"sheet column width:{sheet2.GetColumnWidth(0)}, minWithRow1Only:{minWithRow1Only}");
            Assert.IsTrue(sheet2.GetColumnWidth(0) <= maxWithRow1Only, $"sheet column width:{sheet2.GetColumnWidth(0)}, maxWithRow1Only:{maxWithRow1Only}");

            // Remove the 2nd row merged region and Check that the 2nd row value is used to the AutoSizeColumn width
            sheet2.RemoveMergedRegion(1);
            sheet2.AutoSizeColumn(0);
            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            NPOI.SS.UserModel.ISheet sheet3 = wb3.GetSheet(sheetName);
            Assert.IsTrue(sheet3.GetColumnWidth(0) >= minWithRow1And2);
            Assert.IsTrue(sheet3.GetColumnWidth(0) <= maxWithRow1And2);

            wb3.Close();
            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void TestAutoSizeDate()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet("Sheet1");
            IRow r = s.CreateRow(0);
            r.CreateCell(0).SetCellValue(1);
            r.CreateCell(1).SetCellValue(123456);

            // Will be sized fairly small
            s.AutoSizeColumn((short)0);
            s.AutoSizeColumn((short)1);

            // Size ranges due to different fonts on different machines
            Assert.IsTrue(s.GetColumnWidth(0) > 350, "Single number column too small: " + s.GetColumnWidth(0));
            //Assert.IsTrue(s.GetColumnWidth(0) < 550, "Single number column too big: " + s.GetColumnWidth(0));
            //Todo: find a algorithm of function SheetUtil.GetColumnWidth to make the test statement above succeed.
            Assert.IsTrue(s.GetColumnWidth(0) < 650, "Single number column too big: " + s.GetColumnWidth(0));
            Assert.IsTrue(s.GetColumnWidth(1) > 1500, "6 digit number column too small: " + s.GetColumnWidth(1));
            Assert.IsTrue(s.GetColumnWidth(1) < 2000, "6 digit number column too big: " + s.GetColumnWidth(1));

            // Set a date format
            ICellStyle cs = wb.CreateCellStyle();
            HSSFDataFormat f = (HSSFDataFormat)wb.CreateDataFormat();
            cs.DataFormat = (/*setter*/f.GetFormat("yyyy-mm-dd MMMM hh:mm:ss"));
            r.GetCell(0).CellStyle = (/*setter*/cs);
            r.GetCell(1).CellStyle = (/*setter*/cs);

            Assert.IsTrue(DateUtil.IsCellDateFormatted(r.GetCell(0)));
            Assert.IsTrue(DateUtil.IsCellDateFormatted(r.GetCell(1)));

            // Should Get much bigger now
            s.AutoSizeColumn((short)0);
            s.AutoSizeColumn((short)1);

            Assert.IsTrue(s.GetColumnWidth(0) > 4750, "Date column too small: " + s.GetColumnWidth(0));
            Assert.IsTrue(s.GetColumnWidth(1) > 4750, "Date column too small: " + s.GetColumnWidth(1));
            Assert.IsTrue(s.GetColumnWidth(0) < 6500, "Date column too big: " + s.GetColumnWidth(0));
            Assert.IsTrue(s.GetColumnWidth(0) < 6500, "Date column too big: " + s.GetColumnWidth(0));

            wb.Close();
        }


        [Test]
        public void TestAutoSizeRow()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet 1");

            var row = sheet.CreateRow(0);
            var cell = row.CreateCell(13);
            cell.SetCellValue("test");
            var font = cell.CellStyle.GetFont(workbook);
            font.FontHeightInPoints = 20;
            cell.CellStyle.SetFont(font);
            row.Height = 100;

            sheet.AutoSizeRow(row.RowNum);

            Assert.AreNotEqual(100, row.Height);
            Assert.AreEqual(460, row.Height);

            workbook.Close();
        }


        ///**
        // * Setting ForceFormulaRecalculation on sheets
        // */
        [Test]
        public void TestForceRecalculation()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("UncalcedRecord.xls");

            ISheet sheet = wb1.GetSheetAt(0);
            ISheet sheet2 = wb1.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            row.CreateCell(0).SetCellValue(5);
            row.CreateCell(1).SetCellValue(8);
            Assert.IsFalse(sheet.ForceFormulaRecalculation);
            Assert.IsFalse(sheet2.ForceFormulaRecalculation);

            // Save and manually verify that on column C we have 0, value in template
            HSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();
            sheet.ForceFormulaRecalculation = (/*setter*/true);
            Assert.IsTrue(sheet.ForceFormulaRecalculation);

            // Save and manually verify that on column C we have now 13, calculated value

            // Try it can be opened
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            // And check correct sheet Settings found
            sheet = wb2.GetSheetAt(0);
            sheet2 = wb2.GetSheetAt(1);
            Assert.IsTrue(sheet.ForceFormulaRecalculation);
            Assert.IsFalse(sheet2.ForceFormulaRecalculation);

            // Now turn if back off again
            sheet.ForceFormulaRecalculation = (/*setter*/false);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();

            Assert.IsFalse(wb3.GetSheetAt(0).ForceFormulaRecalculation);
            Assert.IsFalse(wb3.GetSheetAt(1).ForceFormulaRecalculation);
            Assert.IsFalse(wb3.GetSheetAt(2).ForceFormulaRecalculation);

            // Now add a new sheet, and check things work
            //  with old ones unset, new one Set
            ISheet s4 = wb3.CreateSheet();
            s4.ForceFormulaRecalculation = (/*setter*/true);

            Assert.IsFalse(sheet.ForceFormulaRecalculation);
            Assert.IsFalse(sheet2.ForceFormulaRecalculation);
            Assert.IsTrue(s4.ForceFormulaRecalculation);

            HSSFWorkbook wb4 = HSSFTestDataSamples.WriteOutAndReadBack(wb3);
            wb3.Close();

            Assert.IsFalse(wb4.GetSheetAt(0).ForceFormulaRecalculation);
            Assert.IsFalse(wb4.GetSheetAt(1).ForceFormulaRecalculation);
            Assert.IsFalse(wb4.GetSheetAt(2).ForceFormulaRecalculation);
            Assert.IsTrue(wb4.GetSheetAt(3).ForceFormulaRecalculation);

            wb4.Close();
        }

        [Test]
        public new void TestColumnWidth()
        {
            //check we can correctly read column widths from a reference workbook
            IWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("colwidth.xls");

            //reference values
            int[] ref1 = { 365, 548, 731, 914, 1097, 1280, 1462, 1645, 1828, 2011, 2194, 2377, 2560, 2742, 2925, 3108, 3291, 3474, 3657 };

            ISheet sh = wb1.GetSheetAt(0);
            for (char i = 'A'; i <= 'S'; i++)
            {
                int idx = i - 'A';
                int w = sh.GetColumnWidth(idx);
                Assert.AreEqual(ref1[idx], w);
            }

            //the second sheet doesn't have overridden column widths
            sh = wb1.GetSheetAt(1);
            int def_width = sh.DefaultColumnWidth;
            for (char i = 'A'; i <= 'S'; i++)
            {
                int idx = i - 'A';
                int w = sh.GetColumnWidth(idx);
                //getDefaultColumnWidth returns width measured in characters
                //getColumnWidth returns width measured in 1/256th units
                Assert.AreEqual(def_width * 256, w);
            }
            wb1.Close();

            //test new workbook
            HSSFWorkbook wb2 = new HSSFWorkbook();
            sh = wb2.CreateSheet();
            sh.DefaultColumnWidth = (/*setter*/10);
            Assert.AreEqual(10, sh.DefaultColumnWidth);
            Assert.AreEqual(256 * 10, sh.GetColumnWidth(0));
            Assert.AreEqual(256 * 10, sh.GetColumnWidth(1));
            Assert.AreEqual(256 * 10, sh.GetColumnWidth(2));
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                sh.SetColumnWidth(i, w);
                Assert.AreEqual(w, sh.GetColumnWidth(i));
            }

            //serialize and read again
            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();

            sh = wb3.GetSheetAt(0);
            Assert.AreEqual(10, sh.DefaultColumnWidth);
            //columns A-C have default width
            Assert.AreEqual(256 * 10, sh.GetColumnWidth(0));
            Assert.AreEqual(256 * 10, sh.GetColumnWidth(1));
            Assert.AreEqual(256 * 10, sh.GetColumnWidth(2));
            //columns D-F have custom width
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                Assert.AreEqual(w, sh.GetColumnWidth(i));
            }

            // check for 16-bit signed/unsigned error:
            sh.SetColumnWidth(0, 40000);
            Assert.AreEqual(40000, sh.GetColumnWidth(0));

            wb3.Close();
        }
        [Test]
        public void TestDefaultColumnWidth()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("12843-1.xls");
            ISheet sheet = wb1.GetSheetAt(7);
            // shall not be NPE
            Assert.AreEqual(8, sheet.DefaultColumnWidth);
            Assert.AreEqual(8 * 256, sheet.GetColumnWidth(0));

            Assert.AreEqual(0xFF, sheet.DefaultRowHeight);

            wb1.Close();

            HSSFWorkbook wb2 = HSSFTestDataSamples.OpenSampleWorkbook("34775.xls");
            // second and third sheets miss DefaultColWidthRecord
            for (int i = 1; i <= 2; i++)
            {
                int dw = wb2.GetSheetAt(i).DefaultColumnWidth;
                Assert.AreEqual(8, dw);
                int cw = wb2.GetSheetAt(i).GetColumnWidth(0);
                Assert.AreEqual(8 * 256, cw);

                Assert.AreEqual(0xFF, sheet.DefaultRowHeight);
            }

            wb2.Close();
        }
        /**
         * Some utilities Write Excel files without the ROW records.
         * Excel, ooo, and google docs are OK with this.
         * Now POI is too.
         */
        [Test]
        public void TestMissingRowRecords_bug41187()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex41187-19267.xls");

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            Assert.IsNotNull(row, "Identified bug 41187 a");

            Assert.AreNotEqual((short)0, row.Height, "Identified bug 41187 b");

            Assert.AreEqual("Hi Excel!", row.GetCell(0).RichStringCellValue.String);
            // Check row height for 'default' flag
            Assert.AreEqual((short)0xFF, row.Height);

            HSSFTestDataSamples.WriteOutAndReadBack(wb);

            wb.Close();
        }

        /**
         * If we Clone a sheet containing drawings,
         * we must allocate a new ID of the drawing Group and re-Create shape IDs
         *
         * See bug #45720.
         */
        [Test]
        public void TestCloneSheetWithDrawings()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("45720.xls");

            HSSFSheet sheet1 = (HSSFSheet)wb1.GetSheetAt(0);

            wb1.Workbook.FindDrawingGroup();
            DrawingManager2 dm1 = wb1.Workbook.DrawingManager;

            wb1.CloneSheet(0);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            wb2.Workbook.FindDrawingGroup();
            DrawingManager2 dm2 = wb2.Workbook.DrawingManager;

            //Check EscherDggRecord - a workbook-level registry of drawing objects
            Assert.AreEqual(dm1.GetDgg().MaxDrawingGroupId + 1, dm2.GetDgg().MaxDrawingGroupId);

            HSSFSheet sheet2 = (HSSFSheet)wb2.GetSheetAt(1);

            //Check that id of the drawing Group was updated
            EscherDgRecord dg1 = (EscherDgRecord)(sheet1.DrawingPatriarch as HSSFPatriarch).GetBoundAggregate().FindFirstWithId(EscherDgRecord.RECORD_ID);
            EscherDgRecord dg2 = (EscherDgRecord)(sheet2.DrawingPatriarch as HSSFPatriarch).GetBoundAggregate().FindFirstWithId(EscherDgRecord.RECORD_ID);
            int dg_id_1 = dg1.Options >> 4;
            int dg_id_2 = dg2.Options >> 4;
            Assert.AreEqual(dg_id_1 + 1, dg_id_2);

            //TODO: Check shapeId in the Cloned sheet

            wb2.Close();
        }

        /**
         * POI now (Sep 2008) allows sheet names longer than 31 chars (for other apps besides Excel).
         * Since Excel silently truncates to 31, make sure that POI enforces uniqueness on the first
         * 31 chars. 
         */
        [Test]
        public void TestLongSheetNames()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            String SAME_PREFIX = "A123456789B123456789C123456789"; // 30 chars

            try
            {
                wb.CreateSheet(SAME_PREFIX + "Dyyyy"); // identical up to the 32nd char
                Assert.Fail("Expected exception not thrown");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.StartsWith("sheetName 'A123456789B123456789C123456789Dyyyy' is invalid"));
            }
            wb.Close();
        }
        /**
     * Tests that we can read existing column styles
     */
        [Test]
        public void TestReadColumnStyles()
        {
            IWorkbook wbNone = HSSFTestDataSamples.OpenSampleWorkbook("ColumnStyleNone.xls");
            IWorkbook wbSimple = HSSFTestDataSamples.OpenSampleWorkbook("ColumnStyle1dp.xls");
            IWorkbook wbComplex = HSSFTestDataSamples.OpenSampleWorkbook("ColumnStyle1dpColoured.xls");

            // Presence / absence Checks
            Assert.IsNull(wbNone.GetSheetAt(0).GetColumnStyle(0));
            Assert.IsNull(wbNone.GetSheetAt(0).GetColumnStyle(1));

            Assert.IsNull(wbSimple.GetSheetAt(0).GetColumnStyle(0));
            Assert.IsNotNull(wbSimple.GetSheetAt(0).GetColumnStyle(1));

            Assert.IsNull(wbComplex.GetSheetAt(0).GetColumnStyle(0));
            Assert.IsNotNull(wbComplex.GetSheetAt(0).GetColumnStyle(1));

            // Details Checks
            ICellStyle bs = wbSimple.GetSheetAt(0).GetColumnStyle(1);
            Assert.AreEqual(62, bs.Index);
            Assert.AreEqual("#,##0.0_ ;\\-#,##0.0\\ ", bs.GetDataFormatString());
            Assert.AreEqual("Calibri", bs.GetFont(wbSimple).FontName);
            Assert.AreEqual(11 * 20, bs.GetFont(wbSimple).FontHeight);
            Assert.AreEqual(8, bs.GetFont(wbSimple).Color);
            Assert.IsFalse(bs.GetFont(wbSimple).IsItalic);
            Assert.AreEqual((int)FontBoldWeight.Normal, bs.GetFont(wbSimple).Boldweight);


            ICellStyle cs = wbComplex.GetSheetAt(0).GetColumnStyle(1);
            Assert.AreEqual(62, cs.Index);
            Assert.AreEqual("#,##0.0_ ;\\-#,##0.0\\ ", cs.GetDataFormatString());
            Assert.AreEqual("Arial", cs.GetFont(wbComplex).FontName);
            Assert.AreEqual(8 * 20, cs.GetFont(wbComplex).FontHeight);
            Assert.AreEqual(10, cs.GetFont(wbComplex).Color);
            Assert.IsFalse(cs.GetFont(wbComplex).IsItalic);
            Assert.AreEqual((int)FontBoldWeight.Bold, cs.GetFont(wbComplex).Boldweight);

            wbComplex.Close();
            wbSimple.Close();
            wbNone.Close();
        }
        [Test]
        public void TestArabic()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();

            Assert.IsFalse(s.IsRightToLeft);
            s.IsRightToLeft = true;
            Assert.IsTrue(s.IsRightToLeft);

            wb.Close();
        }
        [Test]
        public void TestAutoFilter()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sh = (HSSFSheet)wb1.CreateSheet();
            InternalWorkbook iwb = wb1.Workbook;
            InternalSheet ish = sh.Sheet;

            Assert.IsNull(iwb.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1));
            Assert.IsNull(ish.FindFirstRecordBySid(AutoFilterInfoRecord.sid));

            CellRangeAddress range = CellRangeAddress.ValueOf("A1:B10");
            sh.SetAutoFilter(range);

            NameRecord name = iwb.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1);
            Assert.IsNotNull(name);

            // The built-in name for auto-filter must consist of a single Area3d Ptg.
            Ptg[] ptg = name.NameDefinition;
            Assert.AreEqual(1, ptg.Length, "The built-in name for auto-filter must consist of a single Area3d Ptg");
            Assert.IsTrue(ptg[0] is Area3DPtg, "The built-in name for auto-filter must consist of a single Area3d Ptg");

            Area3DPtg aref = (Area3DPtg)ptg[0];
            Assert.AreEqual(range.FirstColumn, aref.FirstColumn);
            Assert.AreEqual(range.FirstRow, aref.FirstRow);
            Assert.AreEqual(range.LastColumn, aref.LastColumn);
            Assert.AreEqual(range.LastRow, aref.LastRow);

            // verify  AutoFilterInfoRecord
            AutoFilterInfoRecord afilter = (AutoFilterInfoRecord)ish.FindFirstRecordBySid(AutoFilterInfoRecord.sid);
            Assert.IsNotNull(afilter);
            Assert.AreEqual(2, afilter.NumEntries); //filter covers two columns

            HSSFPatriarch dr = (HSSFPatriarch)sh.DrawingPatriarch;
            Assert.IsNotNull(dr);
            HSSFSimpleShape comboBoxShape = (HSSFSimpleShape)dr.Children[0];
            Assert.AreEqual(comboBoxShape.ShapeType, HSSFSimpleShape.OBJECT_TYPE_COMBO_BOX);

            Assert.IsNull(ish.FindFirstRecordBySid(ObjRecord.sid)); // ObjRecord will appear after serializetion

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sh = (HSSFSheet)wb2.GetSheetAt(0);
            ish = sh.Sheet;
            ObjRecord objRecord = (ObjRecord)ish.FindFirstRecordBySid(ObjRecord.sid);
            IList subRecords = objRecord.SubRecords;
            Assert.AreEqual(3, subRecords.Count);
            Assert.IsTrue(subRecords[0] is CommonObjectDataSubRecord);
            Assert.IsTrue(subRecords[1] is FtCblsSubRecord); // must be present, see Bug 51481
            Assert.IsTrue(subRecords[2] is LbsDataSubRecord);

            wb2.Close();
        }
        [Test]
        public void TestGetSetColumnHiddenShort()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet 1");
            sheet.SetColumnHidden((short)2, true);
            Assert.IsTrue(sheet.IsColumnHidden((short)2));

            workbook.Close();
        }
        [Test]
        public void TestColumnWidthShort()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            ISheet sheet = wb1.CreateSheet();

            //default column width measured in characters
            sheet.DefaultColumnWidth = ((short)10);
            Assert.AreEqual(10, sheet.DefaultColumnWidth);
            //columns A-C have default width
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth((short)0));
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth((short)1));
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth((short)2));

            //set custom width for D-F
            for (char i = 'D'; i <= 'F'; i++)
            {
                //Sheet#setColumnWidth accepts the width in units of 1/256th of a character width
                int w = 256 * 12;
                sheet.SetColumnWidth((short)i, w);
                Assert.AreEqual(w, sheet.GetColumnWidth((short)i));
            }
            //reset the default column width, columns A-C change, D-F still have custom width
            sheet.DefaultColumnWidth = ((short)20);
            Assert.AreEqual(20, sheet.DefaultColumnWidth);
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth((short)0));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth((short)1));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth((short)2));
            for (char i = 'D'; i <= 'F'; i++)
            {
                int w = 256 * 12;
                Assert.AreEqual(w, sheet.GetColumnWidth((short)i));
            }

            // check for 16-bit signed/unsigned error:
            sheet.SetColumnWidth((short)10, 40000);
            Assert.AreEqual(40000, sheet.GetColumnWidth((short)10));

            //The maximum column width for an individual cell is 255 characters
            try
            {
                sheet.SetColumnWidth((short)9, 256 * 256);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The maximum column width for an individual cell is 255 characters.", e.Message);
            }

            //serialize and read again
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);

            sheet = wb2.GetSheetAt(0);
            Assert.AreEqual(20, sheet.DefaultColumnWidth);
            //columns A-C have default width
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth((short)0));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth((short)1));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth((short)2));
            //columns D-F have custom width
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                Assert.AreEqual(w, sheet.GetColumnWidth((short)i));
            }
            Assert.AreEqual(40000, sheet.GetColumnWidth((short)10));

            wb2.Close();
        }

        [Test]
        public void TestShowInPane()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            sheet.ShowInPane(2, 3);

            try
            {
                sheet.ShowInPane(int.MaxValue, 3);
                Assert.Fail("Should catch exception here");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Maximum row number is 65535", e.Message);
            }

            wb.Close();
        }
        [Test]
        public void TestDrawingRecords()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            /* TODO: NPE?
            sheet.dumpDrawingRecords(false);
            sheet.dumpDrawingRecords(true);*/
            Assert.IsNull(sheet.DrawingEscherAggregate);

            wb.Close();
        }

        [Test]
        public void TestBug55723b()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            // stored with a special name
            Assert.IsNull(wb.Workbook.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1));

            CellRangeAddress range = CellRangeAddress.ValueOf("A:B");
            IAutoFilter filter = sheet.SetAutoFilter(range);
            Assert.IsNotNull(filter);

            // stored with a special name
            NameRecord record = wb.Workbook.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1);
            Assert.IsNotNull(record);

            wb.Close();
        }

        [Test]
        public void Test58746()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            HSSFSheet first = wb.CreateSheet("first") as HSSFSheet;
            first.CreateRow(0).CreateCell(0).SetCellValue(1);

            HSSFSheet second = wb.CreateSheet("second") as HSSFSheet;
            second.CreateRow(0).CreateCell(0).SetCellValue(2);

            HSSFSheet third = wb.CreateSheet("third") as HSSFSheet;
            HSSFRow row = third.CreateRow(0) as HSSFRow;
            row.CreateCell(0).CellFormula = ("first!A1");
            row.CreateCell(1).CellFormula = ("second!A1");
            // re-order for sheet "third"
            wb.SetSheetOrder("third", 0);

            // verify results
            Assert.AreEqual("third", wb.GetSheetAt(0).SheetName);
            Assert.AreEqual("first", wb.GetSheetAt(1).SheetName);
            Assert.AreEqual("second", wb.GetSheetAt(2).SheetName);

            Assert.AreEqual("first!A1", wb.GetSheetAt(0).GetRow(0).GetCell(0).CellFormula);
            Assert.AreEqual("second!A1", wb.GetSheetAt(0).GetRow(0).GetCell(1).CellFormula);

            wb.Close();
        }

        [Test]
        public void Bug59135()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            wb1.CreateSheet().ProtectSheet("1111.2222.3333.1234");
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            Assert.AreEqual(unchecked((short)0xb86b), (wb2.GetSheetAt(0) as HSSFSheet).Password);
            wb2.Close();

            HSSFWorkbook wb3 = new HSSFWorkbook();
            wb3.CreateSheet().ProtectSheet("1111.2222.3333.12345");
            HSSFWorkbook wb4 = HSSFTestDataSamples.WriteOutAndReadBack(wb3);
            wb3.Close();

            Assert.AreEqual(unchecked((short)0xbecc), (wb4.GetSheetAt(0) as HSSFSheet).Password);
            wb4.Close();
        }
    }
}
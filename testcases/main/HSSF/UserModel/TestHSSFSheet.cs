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
    using NPOI.DDF;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.AutoFilter;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using System.Collections;
    using TestCases.HSSF;
    using TestCases.SS.UserModel;

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
            ClassicAssert.AreEqual(3, nameRecord.SheetNumber);

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

            ClassicAssert.IsTrue(sheet.GridsetRecord.Gridset);
#pragma warning disable 0618 //  warning CS0618: 'NPOI.HSSF.UserModel.HSSFSheet.IsGridsPrinted' is obsolete: 'Please use IsPrintGridlines instead'
            s.IsGridsPrinted = true; // <- this is marked obsolete, but using "s.IsPrintGridlines = true;" makes this test fail 8-(
#pragma warning restore 0618
            ClassicAssert.IsFalse(sheet.GridsetRecord.Gridset);

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

            ClassicAssert.IsFalse(record.VCenter);
            ClassicAssert.IsFalse(s.VerticallyCenter);
            s.VerticallyCenter = (true);
            ClassicAssert.IsTrue(record.VCenter);
            ClassicAssert.IsTrue(s.VerticallyCenter);

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

            ClassicAssert.IsFalse(record.HCenter);
            ClassicAssert.IsFalse(s.HorizontallyCenter);
            s.HorizontallyCenter = (true);
            ClassicAssert.IsTrue(record.HCenter);
            ClassicAssert.IsTrue(s.HorizontallyCenter);

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
            ClassicAssert.IsFalse(record.AlternateExpression);
            ClassicAssert.IsFalse(record.AlternateFormula);
            ClassicAssert.IsFalse(record.Autobreaks);
            ClassicAssert.IsFalse(record.Dialog);
            ClassicAssert.IsFalse(record.DisplayGuts);
            ClassicAssert.IsTrue(record.FitToPage);
            ClassicAssert.IsFalse(record.RowSumsBelow);
            ClassicAssert.IsFalse(record.RowSumsRight);

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
            ClassicAssert.IsFalse(record.AlternateExpression);
            ClassicAssert.IsFalse(record.AlternateFormula);
            ClassicAssert.IsTrue(record.Autobreaks);
            ClassicAssert.IsTrue(record.Dialog);
            ClassicAssert.IsTrue(record.DisplayGuts);
            ClassicAssert.IsFalse(record.FitToPage);
            ClassicAssert.IsTrue(record.RowSumsBelow);
            ClassicAssert.IsTrue(record.RowSumsRight);
            ClassicAssert.IsFalse(s.AlternativeExpression);
            ClassicAssert.IsFalse(s.AlternativeFormula);
            ClassicAssert.IsTrue(s.Autobreaks);
            ClassicAssert.IsTrue(s.Dialog);
            ClassicAssert.IsTrue(s.DisplayGuts);
            ClassicAssert.IsFalse(s.FitToPage);
            ClassicAssert.IsTrue(s.RowSumsBelow);
            ClassicAssert.IsTrue(s.RowSumsRight);

            wb.Close();
        }

        /**
         * Setting landscape and portrait stuff on existing sheets
         */
        [Test]
        public void TestPrintSetupLandscapeExisting()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPageBreaks.xls");

            ClassicAssert.AreEqual(3, wb1.NumberOfSheets);

            NPOI.SS.UserModel.ISheet sheetL = wb1.GetSheetAt(0);
            NPOI.SS.UserModel.ISheet sheetPM = wb1.GetSheetAt(1);
            NPOI.SS.UserModel.ISheet sheetLS = wb1.GetSheetAt(2);

            // Check two aspects of the print setup
            ClassicAssert.IsFalse(sheetL.PrintSetup.Landscape);
            ClassicAssert.IsTrue(sheetPM.PrintSetup.Landscape);
            ClassicAssert.IsTrue(sheetLS.PrintSetup.Landscape);
            ClassicAssert.AreEqual(1, sheetL.PrintSetup.Copies);
            ClassicAssert.AreEqual(1, sheetPM.PrintSetup.Copies);
            ClassicAssert.AreEqual(1, sheetLS.PrintSetup.Copies);

            // Change one on each
            sheetL.PrintSetup.Landscape = (true);
            sheetPM.PrintSetup.Landscape = (false);
            sheetPM.PrintSetup.Copies = ((short)3);

            // Check taken
            ClassicAssert.IsTrue(sheetL.PrintSetup.Landscape);
            ClassicAssert.IsFalse(sheetPM.PrintSetup.Landscape);
            ClassicAssert.IsTrue(sheetLS.PrintSetup.Landscape);
            ClassicAssert.AreEqual(1, sheetL.PrintSetup.Copies);
            ClassicAssert.AreEqual(3, sheetPM.PrintSetup.Copies);
            ClassicAssert.AreEqual(1, sheetLS.PrintSetup.Copies);

            // Save and re-load, and Check still there
            IWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheetL = wb1.GetSheetAt(0);
            sheetPM = wb1.GetSheetAt(1);
            sheetLS = wb1.GetSheetAt(2);

            ClassicAssert.IsTrue(sheetL.PrintSetup.Landscape);
            ClassicAssert.IsFalse(sheetPM.PrintSetup.Landscape);
            ClassicAssert.IsTrue(sheetLS.PrintSetup.Landscape);
            ClassicAssert.AreEqual(1, sheetL.PrintSetup.Copies);
            ClassicAssert.AreEqual(3, sheetPM.PrintSetup.Copies);
            ClassicAssert.AreEqual(1, sheetLS.PrintSetup.Copies);

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

            ClassicAssert.AreEqual(0, r1.OutlineLevel);
            ClassicAssert.AreEqual(0, r2.OutlineLevel);
            ClassicAssert.AreEqual(0, r3.OutlineLevel);
            ClassicAssert.AreEqual(0, r4.OutlineLevel);
            ClassicAssert.AreEqual(0, r5.OutlineLevel);

            s.GroupRow(2, 3);

            ClassicAssert.AreEqual(0, r1.OutlineLevel);
            ClassicAssert.AreEqual(0, r2.OutlineLevel);
            ClassicAssert.AreEqual(1, r3.OutlineLevel);
            ClassicAssert.AreEqual(1, r4.OutlineLevel);
            ClassicAssert.AreEqual(0, r5.OutlineLevel);

            // Save and re-Open
            IWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            s = wb2.GetSheetAt(0);
            r1 = (HSSFRow)s.GetRow(0);
            r2 = (HSSFRow)s.GetRow(1);
            r3 = (HSSFRow)s.GetRow(2);
            r4 = (HSSFRow)s.GetRow(3);
            r5 = (HSSFRow)s.GetRow(4);

            ClassicAssert.AreEqual(0, r1.OutlineLevel);
            ClassicAssert.AreEqual(0, r2.OutlineLevel);
            ClassicAssert.AreEqual(1, r3.OutlineLevel);
            ClassicAssert.AreEqual(1, r4.OutlineLevel);
            ClassicAssert.AreEqual(0, r5.OutlineLevel);

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

            ClassicAssert.AreEqual(0, r1.OutlineLevel);
            ClassicAssert.AreEqual(0, r2.OutlineLevel);
            ClassicAssert.AreEqual(0, r3.OutlineLevel);
            ClassicAssert.AreEqual(0, r4.OutlineLevel);
            ClassicAssert.AreEqual(0, r5.OutlineLevel);
            ClassicAssert.AreEqual(0, r6.OutlineLevel);

            // This used to complain about lacking guts records
            s.GroupRow(2, 4);

            ClassicAssert.AreEqual(0, r1.OutlineLevel);
            ClassicAssert.AreEqual(0, r2.OutlineLevel);
            ClassicAssert.AreEqual(1, r3.OutlineLevel);
            ClassicAssert.AreEqual(1, r4.OutlineLevel);
            ClassicAssert.AreEqual(1, r5.OutlineLevel);
            ClassicAssert.AreEqual(0, r6.OutlineLevel);

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

            ClassicAssert.AreEqual(0, r1.OutlineLevel);
            ClassicAssert.AreEqual(0, r2.OutlineLevel);
            ClassicAssert.AreEqual(1, r3.OutlineLevel);
            ClassicAssert.AreEqual(1, r4.OutlineLevel);
            ClassicAssert.AreEqual(1, r5.OutlineLevel);
            ClassicAssert.AreEqual(0, r6.OutlineLevel);

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
            ClassicAssert.AreSame(p1, p2);
            workbook.Close();
        }
        [Test]
        public void TestGetDrawings()
        {
            HSSFWorkbook wb1c = HSSFTestDataSamples.OpenSampleWorkbook("WithChart.xls");
            HSSFWorkbook wb2c = HSSFTestDataSamples.OpenSampleWorkbook("WithTwoCharts.xls");

            // 1 chart sheet -> data on 1st, chart on 2nd
            ClassicAssert.IsNotNull(((HSSFSheet)wb1c.GetSheetAt(0)).DrawingPatriarch);
            ClassicAssert.IsNotNull(((HSSFSheet)wb1c.GetSheetAt(1)).DrawingPatriarch);
            ClassicAssert.IsFalse((((HSSFSheet)wb1c.GetSheetAt(0)).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            ClassicAssert.IsTrue((((HSSFSheet)wb1c.GetSheetAt(1)).DrawingPatriarch as HSSFPatriarch).ContainsChart());

            // 2 chart sheet -> data on 1st, chart on 2nd+3rd
            ClassicAssert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(0)).DrawingPatriarch);
            ClassicAssert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(1)).DrawingPatriarch);
            ClassicAssert.IsNotNull(((HSSFSheet)wb2c.GetSheetAt(2)).DrawingPatriarch);
            ClassicAssert.IsFalse((((HSSFSheet)wb2c.GetSheetAt(0)).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            ClassicAssert.IsTrue((((HSSFSheet)wb2c.GetSheetAt(1)).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            ClassicAssert.IsTrue((((HSSFSheet)wb2c.GetSheetAt(2)).DrawingPatriarch as HSSFPatriarch).ContainsChart());

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
            ClassicAssert.IsFalse(hssfSheet.ObjectProtect);
            hssfSheet.ProtectSheet(passwordA);
            ClassicAssert.IsTrue(hssfSheet.ObjectProtect);
            ClassicAssert.AreEqual(expectedHashA, hssfSheet.Password);
            ClassicAssert.AreEqual(expectedHashA, hssfSheet.Sheet.ProtectionBlock.PasswordHash);


            // Clone the sheet, and make sure the password hash is preserved
            HSSFSheet sheet2 = (HSSFSheet)workbook.CloneSheet(0);
            ClassicAssert.IsTrue(hssfSheet.ObjectProtect);
            ClassicAssert.AreEqual(expectedHashA, sheet2.Sheet.ProtectionBlock.PasswordHash);

            // change the password on the first sheet
            hssfSheet.ProtectSheet(passwordB);
            ClassicAssert.IsTrue(hssfSheet.ObjectProtect);
            ClassicAssert.AreEqual(expectedHashB, hssfSheet.Sheet.ProtectionBlock.PasswordHash);
            ClassicAssert.AreEqual(expectedHashB, hssfSheet.Password);
            // but the cloned sheet's password should remain unchanged
            ClassicAssert.AreEqual(expectedHashA, sheet2.Password);

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
            //ClassicAssert.IsTrue(pb.IsSheetProtected, "Protection should be on");
            //ClassicAssert.IsTrue(pb.IsObjectProtected, "object Protection should be on");
            //ClassicAssert.IsTrue(pb.IsScenarioProtected, "scenario Protection should be on");
            //ClassicAssert.AreEqual(expected, pb.PasswordHash, "well known value for top secret hash should be " + StringUtil.ToHexString(expected).Substring(4));
            ClassicAssert.IsTrue(s.Protect, "Protection should be on");
            ClassicAssert.IsTrue(s.ObjectProtect, "object Protection should be on");
            ClassicAssert.IsTrue(s.ScenarioProtect, "scenario Protection should be on");
            ClassicAssert.AreEqual(expected, s.Password, "well known value for top secret hash should be " + StringUtil.ToHexString(expected).Substring(4));

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
            ClassicAssert.AreEqual(cls, recs[index].GetType());
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
                throw;
            }
            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rc;
            rc = new RecordInspector.RecordCollector();
            ((HSSFSheet)sheet).Sheet.VisitContainedRecords(rc, 0);
            int nRecsWithProtection = rc.Records.Length;

            sheet.ProtectSheet(null);
            rc = new RecordInspector.RecordCollector();
            ((HSSFSheet)sheet).Sheet.VisitContainedRecords(rc, 0);
            int nRecsWithoutProtection = rc.Records.Length;

            ClassicAssert.AreEqual(4, nRecsWithProtection - nRecsWithoutProtection);

            workbook.Close();
        }

        [Test]
        [Obsolete]
        public void TestZoom()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet();
            ClassicAssert.AreEqual(-1, sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid));
            sheet.SetZoom(75);
            ClassicAssert.IsTrue(sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid) > 0);
            SCLRecord sclRecord = (SCLRecord)sheet.Sheet.FindFirstRecordBySid(SCLRecord.sid);
            ClassicAssert.AreEqual(75, 100 * sclRecord.Numerator / sclRecord.Denominator);

            int sclLoc = sheet.Sheet.FindFirstRecordLocBySid(SCLRecord.sid);
            int window2Loc = sheet.Sheet.FindFirstRecordLocBySid(WindowTwoRecord.sid);
            ClassicAssert.IsTrue(sclLoc == window2Loc + 1);

            // verify limits
            try
            {
                sheet.SetZoom(0);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.AreEqual("Numerator must be greater than 0 and less than 65536", e.Message);
            }
            try
            {
                sheet.SetZoom(65536);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.AreEqual("Numerator must be greater than 0 and less than 65536", e.Message);
            }
            try
            {
                sheet.SetZoom(2, 0);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.AreEqual("Denominator must be greater than 0 and less than 65536", e.Message);
            }
            try
            {
                sheet.SetZoom(2, 65536);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.AreEqual("Denominator must be greater than 0 and less than 65536", e.Message);
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
            ClassicAssert.IsNotNull(sheet);

            ClassicAssert.AreEqual(1, sheet.RowBreaks.Length, "1 row page break");
            ClassicAssert.AreEqual(1, sheet.ColumnBreaks.Length, "1 column page break");

            ClassicAssert.IsTrue(sheet.IsRowBroken(22), "No row page break");
            ClassicAssert.IsTrue(sheet.IsColumnBroken((short)4), "No column page break");

            sheet.SetRowBreak(10);
            sheet.SetColumnBreak((short)13);

            ClassicAssert.AreEqual(2, sheet.RowBreaks.Length, "row breaks number");
            ClassicAssert.AreEqual(2, sheet.ColumnBreaks.Length, "column breaks number");

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);

            ClassicAssert.IsTrue(sheet.IsRowBroken(22), "No row page break");
            ClassicAssert.IsTrue(sheet.IsColumnBroken((short)4), "No column page break");

            ClassicAssert.AreEqual(2, sheet.RowBreaks.Length, "row breaks number");
            ClassicAssert.AreEqual(2, sheet.ColumnBreaks.Length, "column breaks number");

            wb2.Close();
        }
        [Test]
        public void TestDBCSName()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("DBCSSheetName.xls");
            wb.GetSheetAt(1);
            ClassicAssert.AreEqual(wb.GetSheetName(1), "\u090f\u0915", "DBCS Sheet Name 2");
            ClassicAssert.AreEqual(wb.GetSheetName(0), "\u091c\u093e", "DBCS Sheet Name 1");
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
            ClassicAssert.IsNotNull(sheet);

            short toprow = (short)100;
            short leftcol = (short)50;
            sheet.ShowInPane(toprow, leftcol);
            ClassicAssert.AreEqual(toprow, sheet.TopRow, "NPOI.SS.UserModel.Sheet.GetTopRow()");
            ClassicAssert.AreEqual(leftcol, sheet.LeftCol, "NPOI.SS.UserModel.Sheet.GetLeftCol()");

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
            ClassicAssert.AreEqual(r.GetCell(0).NumericCellValue, 1, 1); // sanity Check
            ClassicAssert.IsNotNull(r.GetCell(1));
            ClassicAssert.AreEqual(r.GetCell(1).CellFormula, "A1*2");

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
            for (int i = 3; i < 10; i++)
                sheet.CreateRow(i);

            HSSFTestDataSamples.WriteOutAndReadBack(wb2).Close();

            wb2.Close();
        }
        [Test]
        [Platform("Win")]
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
            ClassicAssert.IsTrue(sheet.GetColumnWidth(0) >= minWithRow1And2, "Column autosized with only one row: wrong width");
            ClassicAssert.IsTrue(sheet.GetColumnWidth(0) <= maxWithRow1And2, "Column autosized with only one row: wrong width");

            //Create a region over the 2nd row and auto size the first column
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            ClassicAssert.IsNotNull(sheet.GetMergedRegion(0));
            sheet.AutoSizeColumn(0);
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);

            // Check that the autoSized column width has ignored the 2nd row
            // because it is included in a merged region (Excel like behavior)
            NPOI.SS.UserModel.ISheet sheet2 = wb2.GetSheet(sheetName);
            ClassicAssert.IsTrue(sheet2.GetColumnWidth(0) >= minWithRow1Only, $"sheet column width:{sheet2.GetColumnWidth(0)}, minWithRow1Only:{minWithRow1Only}");
            ClassicAssert.IsTrue(sheet2.GetColumnWidth(0) <= maxWithRow1Only, $"sheet column width:{sheet2.GetColumnWidth(0)}, maxWithRow1Only:{maxWithRow1Only}");

            // Remove the 2nd row merged region and Check that the 2nd row value is used to the AutoSizeColumn width
            sheet2.RemoveMergedRegion(1);
            sheet2.AutoSizeColumn(0);
            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            NPOI.SS.UserModel.ISheet sheet3 = wb3.GetSheet(sheetName);
            ClassicAssert.IsTrue(sheet3.GetColumnWidth(0) >= minWithRow1And2);
            ClassicAssert.IsTrue(sheet3.GetColumnWidth(0) <= maxWithRow1And2);

            wb3.Close();
            wb2.Close();
            wb1.Close();
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

            ClassicAssert.AreNotEqual(100, row.Height);
            ClassicAssert.AreEqual(540, row.Height);

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
            ClassicAssert.IsFalse(sheet.ForceFormulaRecalculation);
            ClassicAssert.IsFalse(sheet2.ForceFormulaRecalculation);

            // Save and manually verify that on column C we have 0, value in template
            HSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();
            sheet.ForceFormulaRecalculation = (/*setter*/true);
            ClassicAssert.IsTrue(sheet.ForceFormulaRecalculation);

            // Save and manually verify that on column C we have now 13, calculated value

            // Try it can be opened
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            // And check correct sheet Settings found
            sheet = wb2.GetSheetAt(0);
            sheet2 = wb2.GetSheetAt(1);
            ClassicAssert.IsTrue(sheet.ForceFormulaRecalculation);
            ClassicAssert.IsFalse(sheet2.ForceFormulaRecalculation);

            // Now turn if back off again
            sheet.ForceFormulaRecalculation = (/*setter*/false);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();

            ClassicAssert.IsFalse(wb3.GetSheetAt(0).ForceFormulaRecalculation);
            ClassicAssert.IsFalse(wb3.GetSheetAt(1).ForceFormulaRecalculation);
            ClassicAssert.IsFalse(wb3.GetSheetAt(2).ForceFormulaRecalculation);

            // Now add a new sheet, and check things work
            //  with old ones unset, new one Set
            ISheet s4 = wb3.CreateSheet();
            s4.ForceFormulaRecalculation = (/*setter*/true);

            ClassicAssert.IsFalse(sheet.ForceFormulaRecalculation);
            ClassicAssert.IsFalse(sheet2.ForceFormulaRecalculation);
            ClassicAssert.IsTrue(s4.ForceFormulaRecalculation);

            HSSFWorkbook wb4 = HSSFTestDataSamples.WriteOutAndReadBack(wb3);
            wb3.Close();

            ClassicAssert.IsFalse(wb4.GetSheetAt(0).ForceFormulaRecalculation);
            ClassicAssert.IsFalse(wb4.GetSheetAt(1).ForceFormulaRecalculation);
            ClassicAssert.IsFalse(wb4.GetSheetAt(2).ForceFormulaRecalculation);
            ClassicAssert.IsTrue(wb4.GetSheetAt(3).ForceFormulaRecalculation);

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
                double w = sh.GetColumnWidth(idx);
                ClassicAssert.AreEqual(ref1[idx], w);
            }

            //the second sheet doesn't have overridden column widths
            sh = wb1.GetSheetAt(1);
            double def_width = sh.DefaultColumnWidth;
            for (char i = 'A'; i <= 'S'; i++)
            {
                int idx = i - 'A';
                double w = sh.GetColumnWidth(idx);
                //getDefaultColumnWidth returns width measured in characters
                //getColumnWidth returns width measured in 1/256th units
                ClassicAssert.AreEqual(def_width * 256, w);
            }
            wb1.Close();

            //test new workbook
            HSSFWorkbook wb2 = new HSSFWorkbook();
            sh = wb2.CreateSheet();
            sh.DefaultColumnWidth = (/*setter*/10);
            ClassicAssert.AreEqual(10, sh.DefaultColumnWidth);
            ClassicAssert.AreEqual(256 * 10, sh.GetColumnWidth(0));
            ClassicAssert.AreEqual(256 * 10, sh.GetColumnWidth(1));
            ClassicAssert.AreEqual(256 * 10, sh.GetColumnWidth(2));
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                sh.SetColumnWidth(i, w);
                ClassicAssert.AreEqual(w, sh.GetColumnWidth(i));
            }

            //serialize and read again
            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();

            sh = wb3.GetSheetAt(0);
            ClassicAssert.AreEqual(10, sh.DefaultColumnWidth);
            //columns A-C have default width
            ClassicAssert.AreEqual(256 * 10, sh.GetColumnWidth(0));
            ClassicAssert.AreEqual(256 * 10, sh.GetColumnWidth(1));
            ClassicAssert.AreEqual(256 * 10, sh.GetColumnWidth(2));
            //columns D-F have custom width
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                ClassicAssert.AreEqual(w, sh.GetColumnWidth(i));
            }

            // check for 16-bit signed/unsigned error:
            sh.SetColumnWidth(0, 40000);
            ClassicAssert.AreEqual(40000, sh.GetColumnWidth(0));

            wb3.Close();
        }
        [Test]
        public void TestDefaultColumnWidth()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("12843-1.xls");
            ISheet sheet = wb1.GetSheetAt(7);
            // shall not be NPE
            ClassicAssert.AreEqual(8, sheet.DefaultColumnWidth);
            ClassicAssert.AreEqual(8 * 256, sheet.GetColumnWidth(0));

            ClassicAssert.AreEqual(0xFF, sheet.DefaultRowHeight);

            wb1.Close();

            HSSFWorkbook wb2 = HSSFTestDataSamples.OpenSampleWorkbook("34775.xls");
            // second and third sheets miss DefaultColWidthRecord
            for (int i = 1; i <= 2; i++)
            {
                double dw = wb2.GetSheetAt(i).DefaultColumnWidth;
                ClassicAssert.AreEqual(8, dw);
                double cw = wb2.GetSheetAt(i).GetColumnWidth(0);
                ClassicAssert.AreEqual(8 * 256, cw);

                ClassicAssert.AreEqual(0xFF, sheet.DefaultRowHeight);
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
            ClassicAssert.IsNotNull(row, "Identified bug 41187 a");

            ClassicAssert.AreNotEqual((short)0, row.Height, "Identified bug 41187 b");

            ClassicAssert.AreEqual("Hi Excel!", row.GetCell(0).RichStringCellValue.String);
            // Check row height for 'default' flag
            ClassicAssert.AreEqual((short)0xFF, row.Height);

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
            ClassicAssert.AreEqual(dm1.GetDgg().MaxDrawingGroupId + 1, dm2.GetDgg().MaxDrawingGroupId);

            HSSFSheet sheet2 = (HSSFSheet)wb2.GetSheetAt(1);

            //Check that id of the drawing Group was updated
            EscherDgRecord dg1 = (EscherDgRecord)(sheet1.DrawingPatriarch as HSSFPatriarch).GetBoundAggregate().FindFirstWithId(EscherDgRecord.RECORD_ID);
            EscherDgRecord dg2 = (EscherDgRecord)(sheet2.DrawingPatriarch as HSSFPatriarch).GetBoundAggregate().FindFirstWithId(EscherDgRecord.RECORD_ID);
            int dg_id_1 = dg1.Options >> 4;
            int dg_id_2 = dg2.Options >> 4;
            ClassicAssert.AreEqual(dg_id_1 + 1, dg_id_2);

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
                ClassicAssert.IsTrue(e.Message.StartsWith("sheetName 'A123456789B123456789C123456789Dyyyy' is invalid"));
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
            ClassicAssert.IsNull(wbNone.GetSheetAt(0).GetColumnStyle(0));
            ClassicAssert.IsNull(wbNone.GetSheetAt(0).GetColumnStyle(1));

            ClassicAssert.IsNull(wbSimple.GetSheetAt(0).GetColumnStyle(0));
            ClassicAssert.IsNotNull(wbSimple.GetSheetAt(0).GetColumnStyle(1));

            ClassicAssert.IsNull(wbComplex.GetSheetAt(0).GetColumnStyle(0));
            ClassicAssert.IsNotNull(wbComplex.GetSheetAt(0).GetColumnStyle(1));

            // Details Checks
            ICellStyle bs = wbSimple.GetSheetAt(0).GetColumnStyle(1);
            ClassicAssert.AreEqual(62, bs.Index);
            ClassicAssert.AreEqual("#,##0.0_ ;\\-#,##0.0\\ ", bs.GetDataFormatString());
            ClassicAssert.AreEqual("Calibri", bs.GetFont(wbSimple).FontName);
            ClassicAssert.AreEqual(11 * 20, bs.GetFont(wbSimple).FontHeight);
            ClassicAssert.AreEqual(8, bs.GetFont(wbSimple).Color);
            ClassicAssert.IsFalse(bs.GetFont(wbSimple).IsItalic);
            ClassicAssert.IsFalse(bs.GetFont(wbSimple).IsBold);


            ICellStyle cs = wbComplex.GetSheetAt(0).GetColumnStyle(1);
            ClassicAssert.AreEqual(62, cs.Index);
            ClassicAssert.AreEqual("#,##0.0_ ;\\-#,##0.0\\ ", cs.GetDataFormatString());
            ClassicAssert.AreEqual("Arial", cs.GetFont(wbComplex).FontName);
            ClassicAssert.AreEqual(8 * 20, cs.GetFont(wbComplex).FontHeight);
            ClassicAssert.AreEqual(10, cs.GetFont(wbComplex).Color);
            ClassicAssert.IsFalse(cs.GetFont(wbComplex).IsItalic);
            ClassicAssert.IsTrue(cs.GetFont(wbComplex).IsBold);

            wbComplex.Close();
            wbSimple.Close();
            wbNone.Close();
        }
        [Test]
        public void TestArabic()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();

            ClassicAssert.IsFalse(s.IsRightToLeft);
            s.IsRightToLeft = true;
            ClassicAssert.IsTrue(s.IsRightToLeft);

            wb.Close();
        }
        [Test]
        public void TestAutoFilter()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sh = (HSSFSheet)wb1.CreateSheet();
            InternalWorkbook iwb = wb1.Workbook;
            InternalSheet ish = sh.Sheet;

            ClassicAssert.IsNull(iwb.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1));
            ClassicAssert.IsNull(ish.FindFirstRecordBySid(AutoFilterInfoRecord.sid));

            CellRangeAddress range = CellRangeAddress.ValueOf("A1:B10");
            sh.SetAutoFilter(range);

            NameRecord name = iwb.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1);
            ClassicAssert.IsNotNull(name);

            // The built-in name for auto-filter must consist of a single Area3d Ptg.
            Ptg[] ptg = name.NameDefinition;
            ClassicAssert.AreEqual(1, ptg.Length, "The built-in name for auto-filter must consist of a single Area3d Ptg");
            ClassicAssert.IsTrue(ptg[0] is Area3DPtg, "The built-in name for auto-filter must consist of a single Area3d Ptg");

            Area3DPtg aref = (Area3DPtg)ptg[0];
            ClassicAssert.AreEqual(range.FirstColumn, aref.FirstColumn);
            ClassicAssert.AreEqual(range.FirstRow, aref.FirstRow);
            ClassicAssert.AreEqual(range.LastColumn, aref.LastColumn);
            ClassicAssert.AreEqual(range.LastRow, aref.LastRow);

            // verify  AutoFilterInfoRecord
            AutoFilterInfoRecord afilter = (AutoFilterInfoRecord)ish.FindFirstRecordBySid(AutoFilterInfoRecord.sid);
            ClassicAssert.IsNotNull(afilter);
            ClassicAssert.AreEqual(2, afilter.NumEntries); //filter covers two columns

            HSSFPatriarch dr = (HSSFPatriarch)sh.DrawingPatriarch;
            ClassicAssert.IsNotNull(dr);
            HSSFSimpleShape comboBoxShape = (HSSFSimpleShape)dr.Children[0];
            ClassicAssert.AreEqual(comboBoxShape.ShapeType, HSSFSimpleShape.OBJECT_TYPE_COMBO_BOX);

            ClassicAssert.IsNull(ish.FindFirstRecordBySid(ObjRecord.sid)); // ObjRecord will appear after serializetion

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sh = (HSSFSheet)wb2.GetSheetAt(0);
            ish = sh.Sheet;
            ObjRecord objRecord = (ObjRecord)ish.FindFirstRecordBySid(ObjRecord.sid);
            IList subRecords = objRecord.SubRecords;
            ClassicAssert.AreEqual(3, subRecords.Count);
            ClassicAssert.IsTrue(subRecords[0] is CommonObjectDataSubRecord);
            ClassicAssert.IsTrue(subRecords[1] is FtCblsSubRecord); // must be present, see Bug 51481
            ClassicAssert.IsTrue(subRecords[2] is LbsDataSubRecord);

            wb2.Close();
        }
        [Test]
        public void TestGetSetColumnHiddenShort()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet 1");
            sheet.SetColumnHidden((short)2, true);
            ClassicAssert.IsTrue(sheet.IsColumnHidden((short)2));

            workbook.Close();
        }
        [Test]
        public void TestColumnWidthShort()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            ISheet sheet = wb1.CreateSheet();

            //default column width measured in characters
            sheet.DefaultColumnWidth = ((short)10);
            ClassicAssert.AreEqual(10, sheet.DefaultColumnWidth);
            //columns A-C have default width
            ClassicAssert.AreEqual(256 * 10, sheet.GetColumnWidth((short)0));
            ClassicAssert.AreEqual(256 * 10, sheet.GetColumnWidth((short)1));
            ClassicAssert.AreEqual(256 * 10, sheet.GetColumnWidth((short)2));

            //set custom width for D-F
            for (char i = 'D'; i <= 'F'; i++)
            {
                //Sheet#setColumnWidth accepts the width in units of 1/256th of a character width
                int w = 256 * 12;
                sheet.SetColumnWidth((short)i, w);
                ClassicAssert.AreEqual(w, sheet.GetColumnWidth((short)i));
            }
            //reset the default column width, columns A-C change, D-F still have custom width
            sheet.DefaultColumnWidth = ((short)20);
            ClassicAssert.AreEqual(20, sheet.DefaultColumnWidth);
            ClassicAssert.AreEqual(256 * 20, sheet.GetColumnWidth((short)0));
            ClassicAssert.AreEqual(256 * 20, sheet.GetColumnWidth((short)1));
            ClassicAssert.AreEqual(256 * 20, sheet.GetColumnWidth((short)2));
            for (char i = 'D'; i <= 'F'; i++)
            {
                int w = 256 * 12;
                ClassicAssert.AreEqual(w, sheet.GetColumnWidth((short)i));
            }

            // check for 16-bit signed/unsigned error:
            sheet.SetColumnWidth((short)10, 40000);
            ClassicAssert.AreEqual(40000, sheet.GetColumnWidth((short)10));

            //The maximum column width for an individual cell is 255 characters
            try
            {
                sheet.SetColumnWidth((short)9, 256 * 256);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.AreEqual("The maximum column width for an individual cell is 255 characters.", e.Message);
            }

            //serialize and read again
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);

            sheet = wb2.GetSheetAt(0);
            ClassicAssert.AreEqual(20, sheet.DefaultColumnWidth);
            //columns A-C have default width
            ClassicAssert.AreEqual(256 * 20, sheet.GetColumnWidth((short)0));
            ClassicAssert.AreEqual(256 * 20, sheet.GetColumnWidth((short)1));
            ClassicAssert.AreEqual(256 * 20, sheet.GetColumnWidth((short)2));
            //columns D-F have custom width
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                ClassicAssert.AreEqual(w, sheet.GetColumnWidth((short)i));
            }
            ClassicAssert.AreEqual(40000, sheet.GetColumnWidth((short)10));

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
                ClassicAssert.AreEqual("Maximum row number is 65535", e.Message);
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
            ClassicAssert.IsNull(sheet.DrawingEscherAggregate);

            wb.Close();
        }

        [Test]
        public void TestBug55723b()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            // stored with a special name
            ClassicAssert.IsNull(wb.Workbook.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1));

            CellRangeAddress range = CellRangeAddress.ValueOf("A:B");
            IAutoFilter filter = sheet.SetAutoFilter(range);
            ClassicAssert.IsNotNull(filter);

            // stored with a special name
            NameRecord record = wb.Workbook.GetSpecificBuiltinRecord(NameRecord.BUILTIN_FILTER_DB, 1);
            ClassicAssert.IsNotNull(record);

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
            ClassicAssert.AreEqual("third", wb.GetSheetAt(0).SheetName);
            ClassicAssert.AreEqual("first", wb.GetSheetAt(1).SheetName);
            ClassicAssert.AreEqual("second", wb.GetSheetAt(2).SheetName);

            ClassicAssert.AreEqual("first!A1", wb.GetSheetAt(0).GetRow(0).GetCell(0).CellFormula);
            ClassicAssert.AreEqual("second!A1", wb.GetSheetAt(0).GetRow(0).GetCell(1).CellFormula);

            wb.Close();
        }

        [Test]
        public void Bug59135()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            wb1.CreateSheet().ProtectSheet("1111.2222.3333.1234");
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            ClassicAssert.AreEqual(unchecked((short)0xb86b), (wb2.GetSheetAt(0) as HSSFSheet).Password);
            wb2.Close();

            HSSFWorkbook wb3 = new HSSFWorkbook();
            wb3.CreateSheet().ProtectSheet("1111.2222.3333.12345");
            HSSFWorkbook wb4 = HSSFTestDataSamples.WriteOutAndReadBack(wb3);
            wb3.Close();

            ClassicAssert.AreEqual(unchecked((short)0xbecc), (wb4.GetSheetAt(0) as HSSFSheet).Password);
            wb4.Close();
        }
    }
}
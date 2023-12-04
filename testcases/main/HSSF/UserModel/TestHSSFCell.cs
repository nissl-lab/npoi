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
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Util;
    using NPOI.HSSF.UserModel;

    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Record;
    using System.Text;
    using TestCases.SS.UserModel;

    /**
     * Tests various functionity having to do with Cell.  For instance support for
     * paticular datatypes, etc.
     * @author Andrew C. Oliver (andy at superlinksoftware dot com)
     * @author  Dan Sherman (dsherman at isisph.com)
     * @author Alex Jacoby (ajacoby at gmail.com)
     */
    [TestFixture]
    public class TestHSSFCell : BaseTestCell
    {
        public TestHSSFCell()
            : base(HSSFITestDataProvider.Instance)
        {
            // TestSetTypeStringOnFormulaCell and TestToString are depending on the american culture.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }
        private static HSSFWorkbook WriteOutAndReadBack(HSSFWorkbook original)
        {
            return HSSFTestDataSamples.WriteOutAndReadBack(original);
        }

        /**
         * Checks that the recognition of files using 1904 date windowing
         *  is working properly. Conversion of the date is also an issue,
         *  but there's a separate unit Test for that.
         */
        [Test]
        public void TestDateWindowingRead()
        {
            DateTime date = new DateTime(2000, 1, 1);

            // first Check a file with 1900 Date Windowing
            HSSFWorkbook wb = OpenSample("1900DateWindowing.xls");
            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);

            Assert.AreEqual(date, sheet.GetRow(0).GetCell(0).DateCellValue,
                               "Date from file using 1900 Date Windowing");

            // now Check a file with 1904 Date Windowing
            wb = OpenSample("1904DateWindowing.xls");
            sheet = wb.GetSheetAt(0);

            Assert.AreEqual(date, sheet.GetRow(0).GetCell(0).DateCellValue,
                             "Date from file using 1904 Date Windowing");

            wb.Close();
        }

        /**
         * Checks that dates are properly written to both types of files:
         * those with 1900 and 1904 date windowing.  Note that if the
         * previous Test ({@link #TestDateWindowingRead}) Assert.Fails, the
         * results of this Test are meaningless.
         */
        [Test]
        public void TestDateWindowingWrite()
        {
            DateTime date = new DateTime(2000, 1, 1);

            // first Check a file with 1900 Date Windowing
            HSSFWorkbook wb1 = OpenSample("1900DateWindowing.xls");

            SetCell(wb1, 0, 1, date);
            HSSFWorkbook wb2 = WriteOutAndReadBack(wb1);
            wb1.Close();

            Assert.AreEqual(date,
                            ReadCell(wb2, 0, 1), "Date from file using 1900 Date Windowing");

            // now Check a file with 1904 Date Windowing
            wb1 = OpenSample("1904DateWindowing.xls");

            SetCell(wb1, 0, 1, date);
            wb2 = WriteOutAndReadBack(wb1);
            Assert.AreEqual(date,
                            ReadCell(wb2, 0, 1), "Date from file using 1900 Date Windowing");

            wb1.Close();
            wb2.Close();
        }

        /**
 * Test for small bug observable around r736460 (prior to version 3.5).  POI fails to remove
 * the {@link StringRecord} following the {@link FormulaRecord} after the result type had been 
 * changed to number/boolean/error.  Excel silently ignores the extra record, but some POI
 * versions (prior to bug 46213 / r717883) crash instead.
 */
        [Test]
        public void TestCachedTypeChange()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet("Sheet1");
            HSSFCell cell = (HSSFCell)sheet.CreateRow(0).CreateCell(0);
            cell.CellFormula = ("A1");
            cell.SetCellValue("abc");
            ConfirmStringRecord(sheet, true);
            cell.SetCellValue(123);
            NPOI.HSSF.Record.Record[] recs = RecordInspector.GetRecords(sheet, 0);
            if (recs.Length == 28 && recs[23] is StringRecord)
            {
                wb.Close();
                throw new AssertionException("Identified bug - leftover StringRecord");
            }
            ConfirmStringRecord(sheet, false);

            // string to error code
            cell.SetCellValue("abc");
            ConfirmStringRecord(sheet, true);
            cell.SetCellErrorValue(FormulaError.REF);
            ConfirmStringRecord(sheet, false);

            // string to boolean
            cell.SetCellValue("abc");
            ConfirmStringRecord(sheet, true);
            cell.SetCellValue(false);
            ConfirmStringRecord(sheet, false);
            wb.Close();
        }

        private static void ConfirmStringRecord(HSSFSheet sheet, bool isPresent)
        {
            Record[] recs = RecordInspector.GetRecords(sheet, 0);
            Assert.AreEqual(isPresent ? 29 : 28, recs.Length); //for SheetExtRecord
            //Assert.AreEqual(isPresent ? 28 : 27, recs.Length);  // statement in poi, why use above line?
            int index = 22;
            Record fr = recs[index++];
            Assert.AreEqual(typeof(FormulaRecord), fr.GetType());
            if (isPresent)
            {
                Assert.AreEqual(typeof(StringRecord), recs[index++].GetType());
            }
            else
            {
                Assert.IsFalse(typeof(StringRecord) == recs[index].GetType());
            }
            Record dbcr = recs[index];
            Assert.AreEqual(typeof(DBCellRecord), dbcr.GetType());
        }

        private static void SetCell(HSSFWorkbook workbook, int rowIdx, int colIdx, DateTime date)
        {
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(rowIdx);
            ICell cell = row.GetCell(colIdx);

            if (cell == null)
            {
                cell = row.CreateCell(colIdx);
            }
            cell.SetCellValue(date);
        }

        private static DateTime? ReadCell(HSSFWorkbook workbook, int rowIdx, int colIdx)
        {
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(rowIdx);
            ICell cell = row.GetCell(colIdx);
            return cell.DateCellValue;
        }

        /**
         * Tests that the active cell can be correctly read and set
         */
        [Test]
        public void TestActiveCell()
        {
            //read in sample
            HSSFWorkbook wb1 = OpenSample("Simple.xls");

            //Check initial position
            HSSFSheet umSheet = (HSSFSheet)wb1.GetSheetAt(0);
            InternalSheet s = umSheet.Sheet;
            Assert.AreEqual(0, s.ActiveCellCol, "Initial active cell should be in col 0");
            Assert.AreEqual(1, s.ActiveCellRow, "Initial active cell should be on row 1");

            //modify position through Cell
            ICell cell = umSheet.CreateRow(3).CreateCell(2);
            cell.SetAsActiveCell();
            Assert.AreEqual(2, s.ActiveCellCol, "After modify, active cell should be in col 2");
            Assert.AreEqual(3, s.ActiveCellRow, "After modify, active cell should be on row 3");

            //Write book to temp file; read and Verify that position is serialized
            HSSFWorkbook wb2 = WriteOutAndReadBack(wb1);
            wb1.Close();

            umSheet = (HSSFSheet)wb2.GetSheetAt(0);
            s = umSheet.Sheet;

            Assert.AreEqual(2, s.ActiveCellCol, "After serialize, active cell should be in col 2");
            Assert.AreEqual(3, s.ActiveCellRow, "After serialize, active cell should be on row 3");

            wb2.Close();
        }

        [Test]
        public void TestActiveCellBug56114()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sh = wb.CreateSheet();

            sh.CreateRow(0);
            sh.CreateRow(1);
            sh.CreateRow(2);
            sh.CreateRow(3);

            ICell cell = sh.GetRow(1).CreateCell(3);
            sh.GetRow(3).CreateCell(3);

            Assert.AreEqual(0, ((HSSFSheet)wb.GetSheetAt(0)).Sheet.ActiveCellRow);
            Assert.AreEqual(0, ((HSSFSheet)wb.GetSheetAt(0)).Sheet.ActiveCellCol);

            cell.SetAsActiveCell();

            Assert.AreEqual(1, ((HSSFSheet)wb.GetSheetAt(0)).Sheet.ActiveCellRow);
            Assert.AreEqual(3, ((HSSFSheet)wb.GetSheetAt(0)).Sheet.ActiveCellCol);

            //	    FileOutputStream fos = new FileOutputStream("/tmp/56114.xls");
            //
            //	    wb.Write(fos);
            //
            //	    fos.Close();

            IWorkbook wbBack = _testDataProvider.WriteOutAndReadBack(wb);

            Assert.AreEqual(1, ((HSSFSheet)wbBack.GetSheetAt(0)).Sheet.ActiveCellRow);
            Assert.AreEqual(3, ((HSSFSheet)wbBack.GetSheetAt(0)).Sheet.ActiveCellCol);

            wbBack.GetSheetAt(0).GetRow(3).GetCell(3).SetAsActiveCell();

            Assert.AreEqual(3, ((HSSFSheet)wbBack.GetSheetAt(0)).Sheet.ActiveCellRow);
            Assert.AreEqual(3, ((HSSFSheet)wbBack.GetSheetAt(0)).Sheet.ActiveCellCol);

            //	    fos = new FileOutputStream("/tmp/56114a.xls");
            //
            //	    wb.Write(fos);
            //
            //	    fos.Close();

            IWorkbook wbBack2 = _testDataProvider.WriteOutAndReadBack(wbBack);
            wbBack.Close();

            Assert.AreEqual(3, ((HSSFSheet)wbBack2.GetSheetAt(0)).Sheet.ActiveCellRow);
            Assert.AreEqual(3, ((HSSFSheet)wbBack2.GetSheetAt(0)).Sheet.ActiveCellCol);

            wbBack2.Close();
        }


        /**
         * Test reading hyperlinks
         */
        [Test]
        public void TestWithHyperlink()
        {

            HSSFWorkbook wb = OpenSample("WithHyperlink.xls");

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
            ICell cell = sheet.GetRow(4).GetCell(0);
            IHyperlink link = cell.Hyperlink;
            Assert.IsNotNull(link);

            Assert.AreEqual("Foo", link.Label);
            Assert.AreEqual(link.Address, "http://poi.apache.org/");
            Assert.AreEqual(4, link.FirstRow);
            Assert.AreEqual(0, link.FirstColumn);

            wb.Close();
        }

        /**
         * Test reading hyperlinks
         */
        [Test]
        public void TestWithTwoHyperlinks()
        {

            HSSFWorkbook wb = OpenSample("WithTwoHyperLinks.xls");

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);

            ICell cell1 = sheet.GetRow(4).GetCell(0);
            IHyperlink link1 = cell1.Hyperlink;
            Assert.IsNotNull(link1);
            Assert.AreEqual("Foo", link1.Label);
            Assert.AreEqual("http://poi.apache.org/", link1.Address);
            Assert.AreEqual(4, link1.FirstRow);
            Assert.AreEqual(0, link1.FirstColumn);

            ICell cell2 = sheet.GetRow(8).GetCell(1);
            IHyperlink link2 = cell2.Hyperlink;
            Assert.IsNotNull(link2);
            Assert.AreEqual("Bar", link2.Label);
            Assert.AreEqual("http://poi.apache.org/hssf/", link2.Address);
            Assert.AreEqual(8, link2.FirstRow);
            Assert.AreEqual(1, link2.FirstColumn);

            wb.Close();
        }



        [Test]
        public void TestHSSFCellToStringWithDataFormat()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
            cell.SetCellValue(new DateTime(2009, 8, 20));
            NPOI.SS.UserModel.ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy");
            cell.CellStyle = cellStyle;
            Assert.AreEqual("8/20/09", cell.ToString());

            NPOI.SS.UserModel.ICellStyle cellStyle2 = wb.CreateCellStyle();
            IDataFormat format = wb.CreateDataFormat();
            cellStyle2.DataFormat = format.GetFormat("YYYY-mm/dd");
            cell.CellStyle = cellStyle2;
            Assert.AreEqual("2009-08/20", cell.ToString());

            wb.Close();
        }
        [Test]
        public void TestGetDataFormatUniqueIndex()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            IDataFormat format = wb.CreateDataFormat();
            short formatidx1 = format.GetFormat("YYYY-mm/dd");
            short formatidx2 = format.GetFormat("YYYY-mm/dd");
            Assert.AreEqual(formatidx1, formatidx2);
            short formatidx3 = format.GetFormat("000000.000");
            Assert.AreNotEqual(formatidx1, formatidx3);

            wb.Close();
        }
        /**
         * Test to ensure we can only assign cell styles that belong
         *  to our workbook, and not those from other workbooks.
         */
        [Test]
        public void TestCellStyleWorkbookMatch()
        {
            HSSFWorkbook wbA = new HSSFWorkbook();
            HSSFWorkbook wbB = new HSSFWorkbook();

            HSSFCellStyle styA = (HSSFCellStyle)wbA.CreateCellStyle();
            HSSFCellStyle styB = (HSSFCellStyle)wbB.CreateCellStyle();

            styA.VerifyBelongsToWorkbook(wbA);
            styB.VerifyBelongsToWorkbook(wbB);
            try
            {
                styA.VerifyBelongsToWorkbook(wbB);
                Assert.Fail();
            }
            catch (ArgumentException) { }
            try
            {
                styB.VerifyBelongsToWorkbook(wbA);
                Assert.Fail();
            }
            catch (ArgumentException) { }

            ICell cellA = wbA.CreateSheet().CreateRow(0).CreateCell(0);
            ICell cellB = wbB.CreateSheet().CreateRow(0).CreateCell(0);

            cellA.CellStyle = (styA);
            cellB.CellStyle = (styB);
            try
            {
                cellA.CellStyle = (styB);
                Assert.Fail();
            }
            catch (ArgumentException) { }
            try
            {
                cellB.CellStyle = (styA);
                Assert.Fail();
            }
            catch (ArgumentException) { }

            wbA.Close();
            wbB.Close();
        }
        /**
          * HSSF prior to version 3.7 had a bug: it could write a NaN but could not read such a file back.
          */
        [Test]
        public void TestReadNaN()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("49761.xls");
            Assert.IsNotNull(wb);
            wb.Close();
        }

        [Test]
        public void TestHSSFCell1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;
            row.CreateCell(0);
            HSSFCell cell = new HSSFCell(wb, sheet, 0, (short)0);
            Assert.IsNotNull(cell);
            wb.Close();
        }

        [Test]
        public void TestDeprecatedMethods()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;

            // cover some deprecated methods and other smaller stuff...
            Assert.AreEqual(wb.Workbook, cell.BoundWorkbook);

            try
            {
                CellType t = cell.CachedFormulaResultType;
                Assert.Fail("Should catch exception");
            }
            catch (InvalidOperationException)
            {
            }

            try
            {
                Assert.IsNotNull(new HSSFCell(wb, sheet, 0, (short)0, CellType.Error + 1));
                Assert.Fail("Should catch exception");
            }
            catch (Exception)
            {
            }

            cell.RemoveCellComment();
            cell.RemoveCellComment();

            wb.Close();
        }

        [Test]
        public void TestCellType()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;

            cell.SetCellType(CellType.Blank);
            Assert.IsNull(cell.DateCellValue);
            Assert.IsFalse(cell.BooleanCellValue);
            Assert.AreEqual("", cell.ToString());

            cell.SetCellType(CellType.String);
            Assert.AreEqual("", cell.ToString());
            cell.SetCellType(CellType.String);
            cell.SetCellValue(1.2);
            cell.SetCellType(CellType.Numeric);
            Assert.AreEqual("1.2", cell.ToString());
            cell.SetCellType(CellType.Boolean);
            Assert.AreEqual("TRUE", cell.ToString());
            cell.SetCellType(CellType.Boolean);
            cell.SetCellValue("" + FormulaError.VALUE.String);
            cell.SetCellType(CellType.Error);
            Assert.AreEqual("#VALUE!", cell.ToString());
            cell.SetCellType(CellType.Error);
            cell.SetCellType(CellType.Boolean);
            Assert.AreEqual("FALSE", cell.ToString());
            cell.SetCellValue(1.2);
            cell.SetCellType(CellType.Numeric);
            Assert.AreEqual("1.2", cell.ToString());
            cell.SetCellType(CellType.Boolean);
            cell.SetCellType(CellType.String);
            cell.SetCellType(CellType.Error);
            cell.SetCellType(CellType.String);
            cell.SetCellValue(1.2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellType(CellType.String);
            Assert.AreEqual("1.2", cell.ToString());

            cell.SetCellValue((String)null);
            cell.SetCellValue((IRichTextString)null);
            wb.Close();
        }

        [Test]
        public void TestGetDateTimeCellValue()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;
            cell.SetCellValue(new DateTime(2022, 5, 10, 13, 20, 50));
            Assert.IsNotNull(cell.DateCellValue);
            Assert.AreEqual(new DateTime(2022, 5, 10, 13, 20, 50), cell.DateCellValue);
#if NET6_0_OR_GREATER
            Assert.AreEqual(new DateOnly(2022, 5, 10), cell.DateOnlyCellValue);
            Assert.AreEqual(new TimeOnly(13, 20, 50), cell.TimeOnlyCellValue);
#endif
            HSSFCell cell2 = row.CreateCell(1) as HSSFCell;
            cell2.SetCellValue("test");
            Assert.IsNull(cell2.DateCellValue);
        }
    }

}
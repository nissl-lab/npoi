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

namespace TestCases.SS.Format
{
    using System;
    using System.Globalization;
    using System.Text;
    using System.Threading;


    using NPOI.HSSF.UserModel;
    using NPOI.SS.Format;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    using NUnit.Framework;

    [TestFixture]
    public class TestCellFormat
    {
        private static String _255_POUND_SIGNS;
        static TestCellFormat()
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i <= 255; i++)
            {
                sb.Append('#');
            }
            _255_POUND_SIGNS = sb.ToString();
        }

        [Test]
        public void TestSome()
        {
            CellFormat fmt = CellFormat.GetInstance(
                    "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)");
            fmt.Apply(1.1);
        }

        [Test]
        public void TestPositiveFormatHasOnePart()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00");
            CellFormatResult result = fmt.Apply(12.345);
            Assert.AreEqual("12.35", result.Text);
        }

        [Test]
        public void TestNegativeFormatHasOnePart()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00");
            CellFormatResult result = fmt.Apply(-12.345);
            Assert.AreEqual("-12.35", result.Text);
        }

        [Test]
        public void TestZeroFormatHasOnePart()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00");
            CellFormatResult result = fmt.Apply(0.0);
            Assert.AreEqual("0.00", result.Text);
        }

        [Test]
        public void TestPositiveFormatHasPosAndNegParts()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00;-0.00");
            CellFormatResult result = fmt.Apply(12.345);
            Assert.AreEqual("12.35", result.Text);
        }

        [Test]
        public void TestNegativeFormatHasPosAndNegParts()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00;-0.00");
            CellFormatResult result = fmt.Apply(-12.345);
            Assert.AreEqual("-12.35", result.Text);
        }

        [Test]
        public void TestNegativeFormatHasPosAndNegParts2()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00;(0.00)");
            CellFormatResult result = fmt.Apply(-12.345);
            Assert.AreEqual("(12.35)", result.Text);
        }

        [Test]
        public void TestZeroFormatHasPosAndNegParts()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00;-0.00");
            CellFormatResult result = fmt.Apply(0.0);
            Assert.AreEqual("0.00", result.Text);
        }

        [Test]
        public void TestFormatWithThreeSections()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00;-0.00;-");

            Assert.AreEqual("12.35", fmt.Apply(12.345).Text);
            Assert.AreEqual("-12.35", fmt.Apply(-12.345).Text);
            Assert.AreEqual("-", fmt.Apply(0.0).Text);
            Assert.AreEqual("abc", fmt.Apply("abc").Text);
        }

        [Test]
        public void TestFormatWithFourSections()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat fmt = CellFormat.GetInstance("0.00;-0.00;-; @ ");

            Assert.AreEqual("12.35", fmt.Apply(12.345).Text);
            Assert.AreEqual("-12.35", fmt.Apply(-12.345).Text);
            Assert.AreEqual("-", fmt.Apply(0.0).Text);
            Assert.AreEqual(" abc ", fmt.Apply("abc").Text);
        }

        [Test]
        public void TestApplyCellForGeneralFormat()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            ICell cell1 = row.CreateCell(1);
            ICell cell2 = row.CreateCell(2);
            ICell cell3 = row.CreateCell(3);
            ICell cell4 = row.CreateCell(4);

            CellFormat cf = CellFormat.GetInstance("General");

            // case Cell.CELL_TYPE_BLANK
            CellFormatResult result0 = cf.Apply(cell0);
            Assert.AreEqual(string.Empty, result0.Text);

            // case Cell.CELL_TYPE_BOOLEAN
            cell1.SetCellValue(true);
            CellFormatResult result1 = cf.Apply(cell1);
            Assert.AreEqual("TRUE", result1.Text);

            // case Cell.CELL_TYPE_NUMERIC
            cell2.SetCellValue(1.23);
            CellFormatResult result2 = cf.Apply(cell2);
            Assert.AreEqual("1.23", result2.Text);

            cell3.SetCellValue(123.0);
            CellFormatResult result3 = cf.Apply(cell3);
            Assert.AreEqual("123", result3.Text);

            // case Cell.CELL_TYPE_STRING
            cell4.SetCellValue("abc");
            CellFormatResult result4 = cf.Apply(cell4);
            Assert.AreEqual("abc", result4.Text);

            wb.Close();
        }

        [Test]
        public void TestApplyCellForAtFormat()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            ICell cell1 = row.CreateCell(1);
            ICell cell2 = row.CreateCell(2);
            ICell cell3 = row.CreateCell(3);
            ICell cell4 = row.CreateCell(4);

            CellFormat cf = CellFormat.GetInstance("@");

            // case Cell.CELL_TYPE_BLANK
            CellFormatResult result0 = cf.Apply(cell0);
            Assert.AreEqual(string.Empty, result0.Text);

            // case Cell.CELL_TYPE_BOOLEAN
            cell1.SetCellValue(true);
            CellFormatResult result1 = cf.Apply(cell1);
            Assert.AreEqual("TRUE", result1.Text);

            // case Cell.CELL_TYPE_NUMERIC
            cell2.SetCellValue(1.23);
            CellFormatResult result2 = cf.Apply(cell2);
            Assert.AreEqual("1.23", result2.Text);

            cell3.SetCellValue(123.0);
            CellFormatResult result3 = cf.Apply(cell3);
            Assert.AreEqual("123", result3.Text);

            // case Cell.CELL_TYPE_STRING
            cell4.SetCellValue("abc");
            CellFormatResult result4 = cf.Apply(cell4);
            Assert.AreEqual("abc", result4.Text);

            wb.Close();
        }

        [Test]
        public void TestApplyCellForDateFormat()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            ICell cell1 = row.CreateCell(1);

            CellFormat cf = CellFormat.GetInstance("dd/mm/yyyy");

            cell0.SetCellValue(10);
            CellFormatResult result0 = cf.Apply(cell0);
            Assert.AreEqual("10/01/1900", result0.Text);

            cell1.SetCellValue(-1);
            CellFormatResult result1 = cf.Apply(cell1);
            Assert.AreEqual(_255_POUND_SIGNS, result1.Text);

            wb.Close();
        }

        [Test]
        public void TestApplyCellForTimeFormat()
        {
            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("hh:mm");

            cell.SetCellValue(DateUtil.ConvertTime("03:04:05"));
            CellFormatResult result = cf.Apply(cell);
            Assert.AreEqual("03:04", result.Text);

            wb.Close();
        }

        [Test]
        public void TestApplyCellForDateFormatAndNegativeFormat()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            ICell cell1 = row.CreateCell(1);

            CellFormat cf = CellFormat.GetInstance("dd/mm/yyyy;(0)");

            cell0.SetCellValue(10);
            CellFormatResult result0 = cf.Apply(cell0);
            Assert.AreEqual("10/01/1900", result0.Text);

            cell1.SetCellValue(-1);
            CellFormatResult result1 = cf.Apply(cell1);
            Assert.AreEqual("(1)", result1.Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasOnePartAndPartHasCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10", cf.Apply(cell).Text);

            cell.SetCellValue(0.123456789012345);
            Assert.AreEqual("0.123456789", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasTwoPartsFirstHasCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;0.000");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.000", cf.Apply(cell).Text);

            cell.SetCellValue(0.123456789012345);
            Assert.AreEqual("0.123", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0.000", cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual("-10.000", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            cell.SetCellValue("TRUE");
            Assert.AreEqual("TRUE", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasTwoPartsBothHaveCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;[>=10]0.000");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.000", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual(_255_POUND_SIGNS, cf.Apply(cell).Text);

            cell.SetCellValue(-0.123456789012345);
            Assert.AreEqual(_255_POUND_SIGNS, cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual(_255_POUND_SIGNS, cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasThreePartsFirstHasCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;0.000;0.0000");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.0000", cf.Apply(cell).Text);

            cell.SetCellValue(0.123456789012345);
            Assert.AreEqual("0.1235", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0.0000", cf.Apply(cell).Text);

            // Second format part ('0.000') is used for negative numbers
            // so result does not have a minus sign
            cell.SetCellValue(-10);
            Assert.AreEqual("10.000", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasThreePartsFirstTwoHaveCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;[>=10]0.000;0.0000");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.000", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0.0000", cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual("-10.0000", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasThreePartsFirstIsDateFirstTwoHaveCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;[>=10]dd/mm/yyyy;0.0");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10/01/1900", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0.0", cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual("-10.0", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasTwoPartsFirstHasConditionSecondIsGeneral()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;General");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0", cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual("-10", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasThreePartsFirstTwoHaveConditionThirdIsGeneral()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;[>=10]0.000;General");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.000", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0", cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual("-10", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("abc", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasFourPartsFirstHasCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;0.000;0.0000;~~@~~");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.0000", cf.Apply(cell).Text);

            cell.SetCellValue(0.123456789012345);
            Assert.AreEqual("0.1235", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0.0000", cf.Apply(cell).Text);

            // Second format part ('0.000') is used for negative numbers
            // so result does not have a minus sign
            cell.SetCellValue(-10);
            Assert.AreEqual("10.000", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("~~abc~~", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasFourPartsSecondHasCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("0.00;[>=100]0.000;0.0000;~~@~~");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.00", cf.Apply(cell).Text);

            cell.SetCellValue(0.123456789012345);
            Assert.AreEqual("0.12", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0.0000", cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual("-10.0000", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("~~abc~~", cf.Apply(cell).Text);

            cell.SetCellValue(true);
            Assert.AreEqual("~~TRUE~~", cf.Apply(cell).Text);

            wb.Close();
        }

        [Test]
        public void TestApplyFormatHasFourPartsFirstTwoHaveCondition()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[>=100]0.00;[>=10]0.000;0.0000;~~@~~");

            cell.SetCellValue(100);
            Assert.AreEqual("100.00", cf.Apply(cell).Text);

            cell.SetCellValue(10);
            Assert.AreEqual("10.000", cf.Apply(cell).Text);

            cell.SetCellValue(0);
            Assert.AreEqual("0.0000", cf.Apply(cell).Text);

            cell.SetCellValue(-10);
            Assert.AreEqual("-10.0000", cf.Apply(cell).Text);

            cell.SetCellValue("abc");
            Assert.AreEqual("~~abc~~", cf.Apply(cell).Text);

            cell.SetCellValue(true);
            Assert.AreEqual("~~TRUE~~", cf.Apply(cell).Text);

            wb.Close();
        }

        /*
         * Test apply(Object value) with a number as parameter
         */
        [Test]
        public void TestApplyObjectNumber()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            CellFormat cf1 = CellFormat.GetInstance("0.000");

            Assert.AreEqual("1.235", cf1.Apply(1.2345).Text);
            Assert.AreEqual("-1.235", cf1.Apply(-1.2345).Text);

            CellFormat cf2 = CellFormat.GetInstance("0.000;(0.000)");

            Assert.AreEqual("1.235", cf2.Apply(1.2345).Text);
            Assert.AreEqual("(1.235)", cf2.Apply(-1.2345).Text);

            CellFormat cf3 = CellFormat.GetInstance("[>1]0.000;0.0000");

            Assert.AreEqual("1.235", cf3.Apply(1.2345).Text);
            Assert.AreEqual("-1.2345", cf3.Apply(-1.2345).Text);

            CellFormat cf4 = CellFormat.GetInstance("0.000;[>1]0.0000");

            Assert.AreEqual("1.235", cf4.Apply(1.2345).Text);
            Assert.AreEqual(_255_POUND_SIGNS, cf4.Apply(-1.2345).Text);
        }

        /*
         * Test apply(Object value) with a Date as parameter
         */
        [Test]
        public void TestApplyObjectDate()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            CellFormat cf1 = CellFormat.GetInstance("m/d/yyyy");
            DateTime date1 = new SimpleDateFormat("M/d/y").Parse("01/11/2012");
            Assert.AreEqual("1/11/2012", cf1.Apply(date1).Text);
        }

        [Test]
        public void TestApplyCellForDateFormatWithConditions()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("[<1]hh:mm:ss AM/PM;[>=1]dd/mm/yyyy hh:mm:ss AM/PM;General");

            cell.SetCellValue(0.5);
            Assert.AreEqual("12:00:00 PM", cf.Apply(cell).Text);

            cell.SetCellValue(1.5);
            Assert.AreEqual("01/01/1900 12:00:00 PM", cf.Apply(cell).Text);

            cell.SetCellValue(-1);
            Assert.AreEqual(_255_POUND_SIGNS, cf.Apply(cell).Text);

            wb.Close();
        }

        /*
         * Test apply(Object value) with a String as parameter
         */
        [Test]
        public void TestApplyObjectString()
        {
            CellFormat cf = CellFormat.GetInstance("0.00");

            Assert.AreEqual("abc", cf.Apply("abc").Text);
        }

        /*
         * Test apply(Object value) with a Boolean as parameter
         */
        [Test]
        public void TestApplyObjectBoolean()
        {
            CellFormat cf1 = CellFormat.GetInstance("0");
            CellFormat cf2 = CellFormat.GetInstance("General");
            CellFormat cf3 = CellFormat.GetInstance("@");

            Assert.AreEqual("TRUE", cf1.Apply(true).Text);
            Assert.AreEqual("FALSE", cf2.Apply(false).Text);
            Assert.AreEqual("TRUE", cf3.Apply(true).Text);
        }
        [Test]
        public void TestSimpleFractionFormat()
        {
            CellFormat cf1 = CellFormat.GetInstance("# ?/?");
            // Create a workbook, row and cell to test with
            IWorkbook wb = new HSSFWorkbook();

            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(123456.6);
                //System.out.println(cf1.apply(cell).text);
                Assert.AreEqual("123456 3/5", cf1.Apply(cell).Text);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestAccountingFormats()
        {
            char pound = '\u00A3';
            char euro = '\u20AC';

            // Accounting -> 0 decimal places, default currency symbol
            String formatDft = "_-\"$\"* #,##0_-;\\-\"$\"* #,##0_-;_-\"$\"* \"-\"_-;_-@_-";
            // Accounting -> 0 decimal places, US currency symbol
            String formatUS = "_-[$$-409]* #,##0_ ;_-[$$-409]* -#,##0 ;_-[$$-409]* \"-\"_-;_-@_-";
            // Accounting -> 0 decimal places, UK currency symbol
            String formatUK = "_-[$" + pound + "-809]* #,##0_-;\\-[$" + pound + "-809]* #,##0_-;_-[$" + pound + "-809]* \"-\"??_-;_-@_-";
            // French style accounting, euro sign comes after not before
            String formatFR = "_-#,##0* [$" + euro + "-40C]_-;\\-#,##0* [$" + euro + "-40C]_-;_-\"-\"??* [$" + euro + "-40C] _-;_-@_-";

            // Has +ve, -ve and zero rules
            CellFormat cfDft = CellFormat.GetInstance(formatDft);
            CellFormat cfUS = CellFormat.GetInstance(formatUS);
            CellFormat cfUK = CellFormat.GetInstance(formatUK);
            CellFormat cfFR = CellFormat.GetInstance(formatFR);

            // For +ve numbers, should be Space + currency symbol + spaces + whole number with commas + space
            // (Except French, which is mostly reversed...)
            Assert.AreEqual(" $   12 ", cfDft.Apply((12.33)).Text);
            Assert.AreEqual(" $   12 ", cfDft.Apply((12.33)).Text);
            Assert.AreEqual(" $   12 ", cfUS.Apply((12.33)).Text);
            Assert.AreEqual(" " + pound + "   12 ", cfUK.Apply((12.33)).Text);
            Assert.AreEqual(" 12   " + euro + " ", cfFR.Apply((12.33)).Text);

            Assert.AreEqual(" $   16,789 ", cfDft.Apply((16789.2)).Text);
            Assert.AreEqual(" $   16,789 ", cfUS.Apply((16789.2)).Text);
            Assert.AreEqual(" " + pound + "   16,789 ", cfUK.Apply((16789.2)).Text);
            Assert.AreEqual(" 16,789   " + euro + " ", cfFR.Apply((16789.2)).Text);

            // For -ve numbers, gets a bit more complicated...
            Assert.AreEqual("-$   12 ", cfDft.Apply((-12.33)).Text);
            Assert.AreEqual(" $   -12 ", cfUS.Apply((-12.33)).Text);
            Assert.AreEqual("-" + pound + "   12 ", cfUK.Apply((-12.33)).Text);
            Assert.AreEqual("-12   " + euro + " ", cfFR.Apply((-12.33)).Text);

            Assert.AreEqual("-$   16,789 ", cfDft.Apply((-16789.2)).Text);
            Assert.AreEqual(" $   -16,789 ", cfUS.Apply((-16789.2)).Text);
            Assert.AreEqual("-" + pound + "   16,789 ", cfUK.Apply((-16789.2)).Text);
            Assert.AreEqual("-16,789   " + euro + " ", cfFR.Apply((-16789.2)).Text);

            // For zero, should be Space + currency symbol + spaces + Minus + spaces
            Assert.AreEqual(" $   - ", cfDft.Apply((0)).Text);
            // TODO Fix the exception this incorrectly triggers
            //Assert.AreEqual(" $   - ",  cfUS.Apply((0)).Text);
            // TODO Fix these to not have an incorrect bonus 0 on the end 
            //Assert.AreEqual(" "+pound+"   -  ", cfUK.Apply((0)).Text);
            //Assert.AreEqual(" -    "+euro+"  ", cfFR.Apply((0)).Text);
        }

        [Test]
        public void TestThreePartComplexFormat1()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            // verify a rather complex format found e.g. in http://wahl.land-oberoesterreich.gv.at/Downloads/bp10.xls
            CellFormatPart posPart = new CellFormatPart("[$-F400]h:mm:ss\\ AM/PM");
            Assert.IsNotNull(posPart);
            DateTime baseTime = new DateTime(1970, 1, 1, 0, 0, 0);
            double exceldata = DateUtil.GetExcelDate(baseTime.AddMilliseconds(12345));
            //format part 'h', means hour, using a 12-hour clock from 1 to 12.(in excel and .net framework)
            //so the excepted value should be 12:00:12 AM
            Assert.AreEqual("12:00:12 AM", posPart.Apply(baseTime.AddMilliseconds(12345)).Text);

            CellFormatPart negPart = new CellFormatPart("[$-F40]h:mm:ss\\ AM/PM");
            Assert.IsNotNull(negPart);
            Assert.AreEqual("12:00:12 AM", posPart.Apply(baseTime.AddMilliseconds(12345)).Text);
            //Assert.IsNotNull(new CellFormatPart("_-* \"\"??_-;_-@_-"));

            CellFormat instance = CellFormat.GetInstance("[$-F400]h:mm:ss\\ AM/PM;[$-F40]h:mm:ss\\ AM/PM;_-* \"\"??_-;_-@_-");
            Assert.IsNotNull(instance);
            Assert.AreEqual("12:00:12 AM", instance.Apply(baseTime.AddMilliseconds(12345)).Text);
        }

        [Test]
        public void TestThreePartComplexFormat2()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            // verify a rather complex format found e.g. in http://wahl.land-oberoesterreich.gv.at/Downloads/bp10.xls
            CellFormatPart posPart = new CellFormatPart("dd/mm/yyyy");
            Assert.IsNotNull(posPart);
            DateTime baseTime = new DateTime(1970, 1, 1);
            Assert.AreEqual("01/01/1970", posPart.Apply(baseTime.AddMilliseconds(12345)).Text);

            CellFormatPart negPart = new CellFormatPart("dd/mm/yyyy");
            Assert.IsNotNull(negPart);
            Assert.AreEqual("01/01/1970", posPart.Apply(baseTime.AddMilliseconds(12345)).Text);
            //Assert.IsNotNull(new CellFormatPart("_-* \"\"??_-;_-@_-"));

            CellFormat instance = CellFormat.GetInstance("dd/mm/yyyy;dd/mm/yyyy;_-* \"\"??_-;_-@_-");
            Assert.IsNotNull(instance);
            Assert.AreEqual("01/01/1970", instance.Apply(baseTime.AddMilliseconds(12345)).Text);
        }
    }
}
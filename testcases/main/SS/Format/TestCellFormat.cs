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
    using System.Windows.Forms;

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
            Label l = new Label();
            CellFormat fmt = CellFormat.GetInstance(
                    "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)");
            fmt.Apply(l, 1.1);
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
        }

        [Test]
        public void TestApplyLabelCellForGeneralFormat()
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

            Label label0 = new Label();
            Label label1 = new Label();
            Label label2 = new Label();
            Label label3 = new Label();
            Label label4 = new Label();

            // case Cell.CELL_TYPE_BLANK
            CellFormatResult result0 = cf.Apply(label0, cell0);
            Assert.AreEqual(string.Empty, result0.Text);
            Assert.AreEqual(string.Empty, label0.Text);

            // case Cell.CELL_TYPE_BOOLEAN
            cell1.SetCellValue(true);
            CellFormatResult result1 = cf.Apply(label1, cell1);
            Assert.AreEqual("TRUE", result1.Text);
            Assert.AreEqual("TRUE", label1.Text);

            // case Cell.CELL_TYPE_NUMERIC
            cell2.SetCellValue(1.23);
            CellFormatResult result2 = cf.Apply(label2, cell2);
            Assert.AreEqual("1.23", result2.Text);
            Assert.AreEqual("1.23", label2.Text);

            cell3.SetCellValue(123.0);
            CellFormatResult result3 = cf.Apply(label3, cell3);
            Assert.AreEqual("123", result3.Text);
            Assert.AreEqual("123", label3.Text);

            // case Cell.CELL_TYPE_STRING
            cell4.SetCellValue("abc");
            CellFormatResult result4 = cf.Apply(label4, cell4);
            Assert.AreEqual("abc", result4.Text);
            Assert.AreEqual("abc", label4.Text);
        }

        [Test]
        public void TestApplyLabelCellForAtFormat()
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

            Label label0 = new Label();
            Label label1 = new Label();
            Label label2 = new Label();
            Label label3 = new Label();
            Label label4 = new Label();

            // case Cell.CELL_TYPE_BLANK
            CellFormatResult result0 = cf.Apply(label0, cell0);
            Assert.AreEqual(string.Empty, result0.Text);
            Assert.AreEqual(string.Empty, label0.Text);

            // case Cell.CELL_TYPE_BOOLEAN
            cell1.SetCellValue(true);
            CellFormatResult result1 = cf.Apply(label1, cell1);
            Assert.AreEqual("TRUE", result1.Text);
            Assert.AreEqual("TRUE", label1.Text);

            // case Cell.CELL_TYPE_NUMERIC
            cell2.SetCellValue(1.23);
            CellFormatResult result2 = cf.Apply(label2, cell2);
            Assert.AreEqual("1.23", result2.Text);
            Assert.AreEqual("1.23", label2.Text);

            cell3.SetCellValue(123.0);
            CellFormatResult result3 = cf.Apply(label3, cell3);
            Assert.AreEqual("123", result3.Text);
            Assert.AreEqual("123", label3.Text);

            // case Cell.CELL_TYPE_STRING
            cell4.SetCellValue("abc");
            CellFormatResult result4 = cf.Apply(label4, cell4);
            Assert.AreEqual("abc", result4.Text);
            Assert.AreEqual("abc", label4.Text);
        }

        [Test]
        public void TestApplyLabelCellForDateFormat()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            ICell cell1 = row.CreateCell(1);

            CellFormat cf = CellFormat.GetInstance("dd/mm/yyyy");

            Label label0 = new Label();
            Label label1 = new Label();

            cell0.SetCellValue(10);
            CellFormatResult result0 = cf.Apply(label0, cell0);
            Assert.AreEqual("10/01/1900", result0.Text);
            Assert.AreEqual("10/01/1900", label0.Text);

            cell1.SetCellValue(-1);
            CellFormatResult result1 = cf.Apply(label1, cell1);
            Assert.AreEqual(_255_POUND_SIGNS, result1.Text);
            Assert.AreEqual(_255_POUND_SIGNS, label1.Text);
        }

        [Test]
        public void TestApplyLabelCellForTimeFormat()
        {
            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            CellFormat cf = CellFormat.GetInstance("hh:mm");

            Label label = new Label();

            cell.SetCellValue(DateUtil.ConvertTime("03:04:05"));
            CellFormatResult result = cf.Apply(label, cell);
            Assert.AreEqual("03:04", result.Text);
            Assert.AreEqual("03:04", label.Text);
        }

        [Test]
        public void TestApplyLabelCellForDateFormatAndNegativeFormat()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            // Create a workbook, IRow and ICell to test with
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            ICell cell1 = row.CreateCell(1);

            CellFormat cf = CellFormat.GetInstance("dd/mm/yyyy;(0)");

            Label label0 = new Label();
            Label label1 = new Label();

            cell0.SetCellValue(10);
            CellFormatResult result0 = cf.Apply(label0, cell0);
            Assert.AreEqual("10/01/1900", result0.Text);
            Assert.AreEqual("10/01/1900", label0.Text);

            cell1.SetCellValue(-1);
            CellFormatResult result1 = cf.Apply(label1, cell1);
            Assert.AreEqual("(1)", result1.Text);
            Assert.AreEqual("(1)", label1.Text);
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
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(123456.6);
            //System.out.println(cf1.apply(cell).text);
            Assert.AreEqual("123456 3/5", cf1.Apply(cell).Text);
        }
    }
}
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

    using NPOI.HSSF;
    using NPOI.SS.UserModel;

    using NUnit.Framework;
    using System.Collections;
    using TestCases.HSSF;
    using NPOI.SS.Util;
    using NPOI.HSSF.UserModel;
    using System.Collections.Generic;
    using System.Globalization;

    /**
     * Unit Tests for HSSFDataFormatter.java
     *
     * @author James May (james dot may at fmr dot com)
     *
     */
    [TestFixture]
    public class TestHSSFDataFormatter
    {

        private HSSFDataFormatter formatter;
        private HSSFWorkbook wb;

        public TestHSSFDataFormatter()
        {

        }

        [SetUp]
        public void SetUp()
        {
            // One or more test methods depends on the american culture.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            // create the formatter to Test
            formatter = new HSSFDataFormatter();

            // create a workbook to Test with
            wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IDataFormat format = wb.CreateDataFormat();

            // create a row and Put some cells in it
            IRow row = sheet.CreateRow(0);

            // date value for July 8 1901 1:19 PM
            double dateNum = 555.555;
            // date value for July 8 1901 11:23 AM
            double timeNum = 555.47431;

            //valid date formats -- all should have "Jul" in output
            String[] goodDatePatterns = {
            "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
            "mmm/d/yy\\ h:mm PM;@",
            "mmmm/d/yy\\ h:mm;@",
            "mmmm/d;@",
            "mmmm/d/yy;@",
            "mmm/dd/yy;@",
            "[$-409]d\\-mmm;@",
            "[$-409]d\\-mmm\\-yy;@",
            "[$-409]dd\\-mmm\\-yy;@",
            "[$-409]mmm\\-yy;@",
            "[$-409]mmmm\\-yy;@",
            "[$-409]mmmm\\ d\\,\\ yyyy;@",
            "[$-409]mmm/d/yy\\ h:mm:ss;@",
            "[$-409]mmmm/d/yy\\ h:mm:ss am;@",
            "[$-409]mmmmm;@",
            "[$-409]mmmmm\\-yy;@",
            "mmmm/d/yyyy;@",
            "[$-409]d\\-mmm\\-yyyy;@"
        };

            //valid time formats - all should have 11:23 in output
            String[] goodTimePatterns = {
           "HH:MM",
           "HH:MM:SS",
           "HH:MM;HH:MM;HH:MM", 
           // This is fun - blue if positive time,
           //  red if negative time or green for zero!
         "[BLUE]HH:MM;[RED]HH:MM;[GREEN]HH:MM", 
           "yyyy-mm-dd hh:mm",
         "yyyy-mm-dd hh:mm:ss",
        };

            // valid number formats
            String[] goodNumPatterns = {
                "#,##0.0000",
                "#,##0;[Red]#,##0",
                "(#,##0.00_);(#,##0.00)",
                "($#,##0.00_);[Red]($#,##0.00)",
                "$#,##0.00",
                "[$-809]#,##0.00", // international format
                "[$-2]#,##0.00", // international format
                "0000.00000%",
                "0.000E+00",
                "0.00E+00",
                "[BLACK]0.00;[COLOR 5]##.##"
        };

            // invalid date formats -- will throw exception in DecimalFormat ctor
            String[] badNumPatterns = {
                "#,#$'#0.0000",
                "'#','#ABC#0;##,##0",
                "000 '123 4'5'6 000",
                "#''0#0'1#10L16EE"
        };

            // create cells with good date patterns
            for (int i = 0; i < goodDatePatterns.Length; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dateNum);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat(goodDatePatterns[i]));
                cell.CellStyle = (/*setter*/cellStyle);
            }
            row = sheet.CreateRow(1);

            // create cells with time patterns
            for (int i = 0; i < goodTimePatterns.Length; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(timeNum);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat(goodTimePatterns[i]));
                cell.CellStyle = (/*setter*/cellStyle);
            }
            row = sheet.CreateRow(2);

            // create cells with num patterns
            for (int i = 0; i < goodNumPatterns.Length; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(-1234567890.12345);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat(goodNumPatterns[i]));
                cell.CellStyle = (/*setter*/cellStyle);
            }
            row = sheet.CreateRow(3);

            // create cells with bad num patterns
            for (int i = 0; i < badNumPatterns.Length; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(1234567890.12345);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat(badNumPatterns[i]));
                cell.CellStyle = (/*setter*/cellStyle);
            }

            // Built in formats

            { // Zip + 4 format
                row = sheet.CreateRow(4);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(123456789);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat("00000-0000"));
                cell.CellStyle = (/*setter*/cellStyle);
            }

            { // Phone number format
                row = sheet.CreateRow(5);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(5551234567D);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat("[<=9999999]###-####;(###) ###-####"));
                cell.CellStyle = (/*setter*/cellStyle);
            }

            { // SSN format
                row = sheet.CreateRow(6);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(444551234);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat("000-00-0000"));
                cell.CellStyle = (/*setter*/cellStyle);
            }

            { // formula cell
                row = sheet.CreateRow(7);
                ICell cell = row.CreateCell(0);
                cell.SetCellType(CellType.Formula);
                cell.CellFormula = (/*setter*/"SUM(12.25,12.25)/100");
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = (/*setter*/format.GetFormat("##.00%;"));
                cell.CellStyle = (/*setter*/cellStyle);
            }

            { // special cell
                row = sheet.CreateRow(8);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(1234567890.12345);
                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.DataFormat = format.GetFormat("#,##0.00 §â§å§Ò.;-#,##0.00 [$§â§å§Ò.-419]");
                cell.CellStyle = (/*setter*/cellStyle);
            }
        }

        /**
         * Test Getting formatted values from numeric and date cells.
         */
        [Test]
        public void TestGetFormattedCellValueHSSFCell()
        {
            // Valid date formats -- cell values should be date formatted & not "555.555"
            IRow row = wb.GetSheetAt(0).GetRow(0);
            IEnumerator it = row.GetEnumerator();
            log("==== VALID DATE FORMATS ====");
            while (it.MoveNext())
            {
                ICell cell = (ICell)it.Current;
                String fmtval = formatter.FormatCellValue(cell);
                log(fmtval);

                // should not be equal to "555.555"
                Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));
                Assert.IsTrue(!"555.555".Equals(fmtval));

                String fmt = cell.CellStyle.GetDataFormatString();

                //assert the correct month form, as in the original Excel format
                String monthPtrn = fmt.IndexOf("mmmm") != -1 ? "MMMM" : "MMM";
                // this line is intended to compute how "July" would look like in the current locale
                String jul = new SimpleDateFormat(monthPtrn).Format(new DateTime(2010, 7, 15), CultureInfo.CurrentCulture);
                // special case for MMMMM = 1st letter of month name
                if (fmt.IndexOf("mmmmm") > -1)
                {
                    jul = jul.Substring(0, 1);
                }
                // check we found july properly
                Assert.IsTrue(fmtval.IndexOf(jul) > -1, "Format came out incorrect - " + fmt);
            }

            row = wb.GetSheetAt(0).GetRow(1);
            it = row.GetEnumerator();
            log("==== VALID TIME FORMATS ====");
            while (it.MoveNext())
            {
                ICell cell = (ICell)it.Current;
                String fmt = cell.CellStyle.GetDataFormatString();
                String fmtval = formatter.FormatCellValue(cell);
                log(fmtval);

                // should not be equal to "555.47431"
                Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));
                Assert.IsTrue(!"555.47431".Equals(fmtval));

                // check we found the time properly
                Assert.IsTrue(fmtval.IndexOf("11:23") > -1, "Format came out incorrect - " + fmt);
            }

            // Test number formats
            row = wb.GetSheetAt(0).GetRow(1);
            it = row.GetEnumerator();
            log("\n==== VALID NUMBER FORMATS ====");
            while (it.MoveNext())
            {
                ICell cell = (ICell)it.Current;
                log(formatter.FormatCellValue(cell));

                // should not be equal to "1234567890.12345"
                Assert.IsTrue(!"1234567890.12345".Equals(formatter.FormatCellValue(cell)));
            }

            // Test bad number formats
            row = wb.GetSheetAt(0).GetRow(3);
            it = row.GetEnumerator();
            log("\n==== INVALID NUMBER FORMATS ====");
            while (it.MoveNext())
            {
                ICell cell = (ICell)it.Current;
                log(formatter.FormatCellValue(cell));
                // should be equal to "1234567890.12345" 
                // in some locales the the decimal delimiter is a comma, not a dot
                string decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                Assert.AreEqual("1234567890" + decimalSeparator + "12345", formatter.FormatCellValue(cell));
            }

            // Test Zip+4 format
            row = wb.GetSheetAt(0).GetRow(4);
            ICell cell1 = row.GetCell(0);
            log("\n==== ZIP FORMAT ====");
            log(formatter.FormatCellValue(cell1));
            Assert.AreEqual("12345-6789", formatter.FormatCellValue(cell1));

            // Test phone number format
            row = wb.GetSheetAt(0).GetRow(5);
            cell1 = row.GetCell(0);
            log("\n==== PHONE FORMAT ====");
            log(formatter.FormatCellValue(cell1));
            Assert.AreEqual("(555) 123-4567", formatter.FormatCellValue(cell1));

            // Test SSN format
            row = wb.GetSheetAt(0).GetRow(6);
            cell1 = row.GetCell(0);
            log("\n==== SSN FORMAT ====");
            log(formatter.FormatCellValue(cell1));
            Assert.AreEqual("444-55-1234", formatter.FormatCellValue(cell1));

            // null Test-- null cell should result in empty String
            Assert.AreEqual(formatter.FormatCellValue(null), "");

            // null Test-- null cell should result in empty String
            Assert.AreEqual(formatter.FormatCellValue(null), "");
        }
        [Test]
        public void TestGetFormattedCellValueHSSFCellHSSFFormulaEvaluator()
        {
            // Test formula format
            IRow row = wb.GetSheetAt(0).GetRow(7);
            ICell cell = row.GetCell(0);
            log("\n==== FORMULA CELL ====");

            // first without a formula Evaluator
            log(formatter.FormatCellValue(cell) + "\t (without Evaluator)");
            Assert.AreEqual("SUM(12.25,12.25)/100", formatter.FormatCellValue(cell));

            // now with a formula Evaluator
            HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(wb);
            log(formatter.FormatCellValue(cell, Evaluator) + "\t\t\t (with Evaluator)");
            string decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            Assert.AreEqual(formatter.FormatCellValue(cell, Evaluator), "24" + decimalSeparator + "50%");

        }

        /**
         * Test using a default number format. The format should be used when a
         * format pattern cannot be Parsed by DecimalFormat.
         */
        [Test]
        public void TestSetDefaultNumberFormat()
        {
            IRow row = wb.GetSheetAt(0).GetRow(3);
            List<ICell> cells = row.Cells;
            FormatBase defaultFormat = new DecimalFormat("Balance $#,#00.00 USD;Balance -$#,#00.00 USD");
            formatter.SetDefaultNumberFormat(defaultFormat);
            Random rand = new Random((int)DateTime.Now.ToFileTime());
            log("\n==== DEFAULT NUMBER FORMAT ====");
            foreach(ICell cell in cells)
            {
                cell.SetCellValue(cell.NumericCellValue * rand.Next() / 1000000 - 1000);
                log(formatter.FormatCellValue(cell));
                Assert.IsTrue(formatter.FormatCellValue(cell).StartsWith("Balance "));
                Assert.IsTrue(formatter.FormatCellValue(cell).EndsWith(" USD"));
            }
        }

        /**
         * A format of "@" means use the general format
         */
        [Test]
        public void TestGeneralAtFormat()
        {
            IWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("47154.xls");
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(0);
            ICell cellA1 = row.GetCell(0);

            Assert.AreEqual(CellType.Numeric, cellA1.CellType);
            Assert.AreEqual(2345.0, cellA1.NumericCellValue, 0.0001);
            Assert.AreEqual("@", cellA1.CellStyle.GetDataFormatString());

            HSSFDataFormatter f = new HSSFDataFormatter();

            Assert.AreEqual("2345", f.FormatCellValue(cellA1));
        }

        /**
         * Tests various formattings of dates and numbers
         */
        [Test]
        public void TestFromFile()
        {
            IWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("Formatting.xls");
            ISheet sheet = workbook.GetSheetAt(0);

            HSSFDataFormatter f = new HSSFDataFormatter();

            // This one is one of the nasty auto-locale changing ones...
            Assert.AreEqual("dd/mm/yyyy", sheet.GetRow(1).GetCell(0).StringCellValue);
            Assert.AreEqual("m/d/yy", sheet.GetRow(1).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("11/24/06", f.FormatCellValue(sheet.GetRow(1).GetCell(1)));

            Assert.AreEqual("yyyy/mm/dd", sheet.GetRow(2).GetCell(0).StringCellValue);
            Assert.AreEqual("yyyy/mm/dd", sheet.GetRow(2).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("2006/11/24", f.FormatCellValue(sheet.GetRow(2).GetCell(1)));

            Assert.AreEqual("yyyy-mm-dd", sheet.GetRow(3).GetCell(0).StringCellValue);
            Assert.AreEqual("yyyy\\-mm\\-dd", sheet.GetRow(3).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("2006-11-24", f.FormatCellValue(sheet.GetRow(3).GetCell(1)));

            Assert.AreEqual("yy/mm/dd", sheet.GetRow(4).GetCell(0).StringCellValue);
            Assert.AreEqual("yy/mm/dd", sheet.GetRow(4).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("06/11/24", f.FormatCellValue(sheet.GetRow(4).GetCell(1)));

            // Another builtin fun one
            Assert.AreEqual("dd/mm/yy", sheet.GetRow(5).GetCell(0).StringCellValue);
            Assert.AreEqual("d/m/yy;@", sheet.GetRow(5).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("24/11/06", f.FormatCellValue(sheet.GetRow(5).GetCell(1)));

            Assert.AreEqual("dd-mm-yy", sheet.GetRow(6).GetCell(0).StringCellValue);
            Assert.AreEqual("dd\\-mm\\-yy", sheet.GetRow(6).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("24-11-06", f.FormatCellValue(sheet.GetRow(6).GetCell(1)));


            // Another builtin fun one
            Assert.AreEqual("nn.nn", sheet.GetRow(9).GetCell(0).StringCellValue);
            Assert.AreEqual("General", sheet.GetRow(9).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("10.52", f.FormatCellValue(sheet.GetRow(9).GetCell(1)));

            // text isn't quite the format rule...
            Assert.AreEqual("nn.nnn", sheet.GetRow(10).GetCell(0).StringCellValue);
            Assert.AreEqual("0.000", sheet.GetRow(10).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("10.520", f.FormatCellValue(sheet.GetRow(10).GetCell(1)));

            // text isn't quite the format rule...
            Assert.AreEqual("nn.n", sheet.GetRow(11).GetCell(0).StringCellValue);
            Assert.AreEqual("0.0", sheet.GetRow(11).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("10.5", f.FormatCellValue(sheet.GetRow(11).GetCell(1)));

            // text isn't quite the format rule...
            Assert.AreEqual("\u00a3nn.nn", sheet.GetRow(12).GetCell(0).StringCellValue);
            Assert.AreEqual("\"\u00a3\"#,##0.00", sheet.GetRow(12).GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual("\u00a310.52", f.FormatCellValue(sheet.GetRow(12).GetCell(1)));
        }

        private static void log(String msg)
        {
            //if (false) { // successful Tests should be silent
            Console.WriteLine(msg);
            //}
        }
    }

}
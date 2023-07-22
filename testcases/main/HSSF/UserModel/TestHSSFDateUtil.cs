
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

    using TestCases.HSSF;
    using NPOI.HSSF.Model;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * Class TestHSSFDateUtil
     *
     *
     * @author  Dan Sherman (dsherman at isisph.com)
     * @author  Hack Kampbjorn (hak at 2mba.dk)
     * @author  Pavel Krupets (pkrupets at palmtreebusiness dot com)
     * @author Alex Jacoby (ajacoby at gmail.com)
     * @version %I%, %G%
     */
    [TestFixture]
    public class TestHSSFDateUtil
    {

        public static int CALENDAR_JANUARY = 1;
        public static int CALENDAR_FEBRUARY = 2;
        public static int CALENDAR_MARCH = 3;
        public static int CALENDAR_APRIL = 4;
        public static int CALENDAR_JULY = 7;
        public static int CALENDAR_OCTOBER = 10;

        /**
         * Checks the date conversion functions in the DateUtil class.
         */
        [Test]
        public void TestDateConversion()
        {

            // Iteratating over the hours exposes any rounding issues.
            for (int hour = 1; hour < 24; hour++)
            {
                DateTime date = new DateTime(2002, 1, 1,
                        hour, 1, 1);
                double excelDate =
                        DateUtil.GetExcelDate(date, false);

                Assert.AreEqual(date,DateUtil.GetJavaDate(excelDate, false), "Checking hour = " + hour);
            }

            // Check 1900 and 1904 date windowing conversions
            double excelDate2 = 36526.0;
            // with 1900 windowing, excelDate is Jan. 1, 2000
            // with 1904 windowing, excelDate is Jan. 2, 2004
            
            DateTime dateIf1900 = new DateTime(2000, 1, 1); // Jan. 1, 2000

            DateTime dateIf1904 = dateIf1900.AddYears(4); // now Jan. 1, 2004
            dateIf1904 = dateIf1904.AddDays(1); // now Jan. 2, 2004
            
            // 1900 windowing
            Assert.AreEqual(dateIf1900,
                    DateUtil.GetJavaDate(excelDate2, false), "Checking 1900 Date Windowing");
            // 1904 windowing
            Assert.AreEqual(
                    dateIf1904,DateUtil.GetJavaDate(excelDate2, true),"Checking 1904 Date Windowing");
        }

        /**
         * Checks the conversion of a java.util.date to Excel on a day when
         * Daylight Saving Time starts.
         */
        
        public void TestExcelConversionOnDSTStart()
        {
            //TODO:: change time zone
            DateTime cal = new DateTime(2004, CALENDAR_MARCH, 28);
            for (int hour = 0; hour < 24; hour++)
            {

                // Skip 02:00 CET as that is the Daylight change time
                // and Java converts it automatically to 03:00 CEST
                if (hour == 2)
                {
                    continue;
                }

                cal.AddHours(hour);
                DateTime javaDate = cal;
                double excelDate = DateUtil.GetExcelDate(javaDate, false);
                double difference = excelDate - Math.Floor(excelDate);
                int differenceInHours = (int)(difference * 24 * 60 + 0.5) / 60;
                Assert.AreEqual(
                        hour,
                        differenceInHours, "Checking " + hour + " hour on Daylight Saving Time start date");
                Assert.AreEqual(
                        javaDate,
                        DateUtil.GetJavaDate(excelDate, false),
                        "Checking " + hour + " hour on Daylight Saving Time start date");
            }
        }

        /**
         * Checks the conversion of an Excel date to a java.util.date on a day when
         * Daylight Saving Time starts.
         */
        
        public void TestJavaConversionOnDSTStart()
        {
            //TODO:: change time zone
            DateTime cal = new DateTime(2004, CALENDAR_MARCH, 28);
            double excelDate = DateUtil.GetExcelDate(cal, false);
            double oneHour = 1.0 / 24;
            double oneMinute = oneHour / 60;
            for (int hour = 0; hour < 24; hour++, excelDate += oneHour)
            {

                // Skip 02:00 CET as that is the Daylight change time
                // and Java converts it automatically to 03:00 CEST
                if (hour == 2)
                {
                    continue;
                }

                cal.AddHours(hour);
                DateTime javaDate = DateUtil.GetJavaDate(excelDate, false);
                Assert.AreEqual(
                        excelDate,
                        DateUtil.GetExcelDate(javaDate, false),0.001);
            }
        }

        /**
         * Checks the conversion of a java.util.Date to Excel on a day when
         * Daylight Saving Time ends.
         */
        
        public void TestExcelConversionOnDSTEnd()
        {
            //TODO:: change time zone
            DateTime cal = new DateTime(2004, CALENDAR_OCTOBER, 31);
            for (int hour = 0; hour < 24; hour++)
            {
                cal.AddDays(hour);
                DateTime javaDate = cal;
                double excelDate = DateUtil.GetExcelDate(javaDate, false);
                double difference = excelDate - Math.Floor(excelDate);
                int differenceInHours = (int)(difference * 24 * 60 + 0.5) / 60;
                Assert.AreEqual(
                        hour,
                        differenceInHours, "Checking " + hour + " hour on Daylight Saving Time end date");
                Assert.AreEqual(
                        javaDate,
                        DateUtil.GetJavaDate(excelDate, false),
                        "Checking " + hour + " hour on Daylight Saving Time start date");
            }
        }

        /**
         * Checks the conversion of an Excel date to java.util.Date on a day when
         * Daylight Saving Time ends.
         */
        [Test]
        public void TestJavaConversionOnDSTEnd()
        {
            //TODO:: change time zone
            DateTime cal = new DateTime(2004, CALENDAR_OCTOBER, 31);
            double excelDate = DateUtil.GetExcelDate(cal, false);
            double oneHour = 1.0 / 24;
            double oneMinute = oneHour / 60;
            for (int hour = 0; hour < 24; hour++, excelDate += oneHour)
            {
                cal.AddHours( hour);
                DateTime javaDate = DateUtil.GetJavaDate(excelDate, false);
                Assert.AreEqual(
                        excelDate,
                        DateUtil.GetExcelDate(javaDate, false),0.1);
            }
        }

        /**
         * Tests that we deal with time-zones properly
         */
        [Ignore("always fail")]
        public void TestCalendarConversion()
        {
            DateTime date = new DateTime(2002, 1, 1, 12, 1, 1);
            DateTime expected = date;

            // Iterating over the hours exposes any rounding issues.
            for (int hour = -12; hour <= 12; hour++)
            {
                String id = "GMT" + (hour < 0 ? "" : "+") + hour + ":00";

                //TODO:: change time zone
                //date.SetTimeZone(TimeZone.GetTimeZone(id));
                //date.AddDays(12);
                
                double excelDate = DateUtil.GetExcelDate(date, false);
                DateTime javaDate = DateUtil.GetJavaDate(excelDate);

                // Should Match despite time-zone
                Assert.AreEqual(expected, javaDate, "Checking timezone " + id);
            }


            //// Check that the timezone aware getter works correctly 
            //TimeZone cet = TimeZone.getTimeZone("Europe/Copenhagen");
            //TimeZone ldn = TimeZone.getTimeZone("Europe/London");
            //TimeZone.setDefault(cet);

            //// 12:45 on 27th April 2012
            //double excelDate = 41026.53125;

            //// Same, no change
            //assertEquals(
            //      HSSFDateUtil.getJavaDate(excelDate, false).getTime(),
            //      HSSFDateUtil.getJavaDate(excelDate, false, cet).getTime()
            //);

            //// London vs Copenhagen, should differ by an hour
            //Date cetDate = HSSFDateUtil.getJavaDate(excelDate, false);
            //Date ldnDate = HSSFDateUtil.getJavaDate(excelDate, false, ldn);
            //assertEquals(ldnDate.getTime() - cetDate.getTime(), 60 * 60 * 1000);
        }

        /**
         * Tests that we correctly detect date formats as such
         */
        [Test]
        public void TestIdentifyDateFormats()
        {
            // First up, try with a few built in date formats
            short[] builtins = new short[] { 0x0e, 0x0f, 0x10, 0x16, 0x2d, 0x2e };
            for (int i = 0; i < builtins.Length; i++)
            {
                String formatStr = HSSFDataFormat.GetBuiltinFormat(builtins[i]);
                Assert.IsTrue(DateUtil.IsInternalDateFormat(builtins[i]));
                Assert.IsTrue(DateUtil.IsADateFormat(builtins[i], formatStr));
            }

            // Now try a few built-in non date formats
            builtins = new short[] { 0x01, 0x02, 0x17, 0x1f, 0x30 };
            for (int i = 0; i < builtins.Length; i++)
            {
                String formatStr = HSSFDataFormat.GetBuiltinFormat(builtins[i]);
                Assert.IsFalse(DateUtil.IsInternalDateFormat(builtins[i]));
                Assert.IsFalse(DateUtil.IsADateFormat(builtins[i], formatStr));
            }

            // Now for some non-internal ones
            // These come after the real ones
            int numBuiltins = HSSFDataFormat.NumberOfBuiltinBuiltinFormats;
            Assert.IsTrue(numBuiltins < 60);
            short formatId = 60;
            Assert.IsFalse(DateUtil.IsInternalDateFormat(formatId));

            // Valid ones first
            String[] formats = new String[] {
                "yyyy-mm-dd", "yyyy/mm/dd", "yy/mm/dd", "yy/mmm/dd",
                "dd/mm/yy", "dd/mm/yyyy", "dd/mmm/yy",
                "dd-mm-yy", "dd-mm-yyyy",
                "DD-MM-YY", "DD-mm-YYYY",
                "dd\\-mm\\-yy", // Sometimes escaped
                "dd.mm.yyyy", "dd\\.mm\\.yyyy",
                "dd\\ mm\\.yyyy AM", "dd\\ mm\\.yyyy pm",
                 "dd\\ mm\\.yyyy\\-dd", "[h]:mm:ss",
                 //"mm/dd/yy", "\"mm\"/\"dd\"/\"yy\"",
                 "mm/dd/yy", 
                 "\\\"mm\\\"/\\\"dd\\\"/\\\"yy\\\"",
                 "mm/dd/yy \"\\\"some\\\" string\"",
                 "m\\/d\\/yyyy", 

                // These crazy ones are valid
                "yyyy-mm-dd;@", "yyyy/mm/dd;@",
                "dd-mm-yy;@", "dd-mm-yyyy;@",
                // These even crazier ones are also valid
                // (who knows what they mean though...)
                "[$-F800]dddd\\,\\ mmm\\ dd\\,\\ yyyy",
                "[$-F900]ddd/mm/yyy",
                // These ones specify colours, who knew that was allowed?
                "[BLACK]dddd/mm/yy",
                "[yeLLow]yyyy-mm-dd"
        };
            for (int i = 0; i < formats.Length; i++)
            {
                Assert.IsTrue(
                        DateUtil.IsADateFormat(formatId, formats[i])
                        ,formats[i] + " is a date format"
                );
            }

            // Then time based ones too
            formats = new String[] {
                "yyyy-mm-dd hh:mm:ss", "yyyy/mm/dd HH:MM:SS",
                "mm/dd HH:MM", "yy/mmm/dd SS",
                "mm/dd HH:MM AM", "mm/dd HH:MM am",
                "mm/dd HH:MM PM", "mm/dd HH:MM pm"
                //"m/d/yy h:mm AM/PM"
        };
            for (int i = 0; i < formats.Length; i++)
            {
                Assert.IsTrue(
                        DateUtil.IsADateFormat(formatId, formats[i]),
                        formats[i] + " is a datetime format"
                );
            }

            // Then invalid ones
            formats = new String[] {
                "yyyy*mm*dd",
                "0.0", "0.000",
                "0%", "0.0%",
                "[]Foo", "[BLACK]0.00%",
                "", null
        };
            for (int i = 0; i < formats.Length; i++)
            {
                Assert.IsFalse(

                        DateUtil.IsADateFormat(formatId, formats[i]), 
                        formats[i] + " is not a date or datetime format"
                );
            }

            // And these are ones we probably shouldn't allow,
            //  but would need a better regexp
            formats = new String[] {
                //"yyyy:mm:dd",
                "yyyy:mm:dd", "\"mm\"/\"dd\"/\"yy\"",
        };
            for (int i = 0; i < formats.Length; i++)
            {
                //    Assert.IsFalse( DateUtil.IsADateFormat(formatId, formats[i]) );
            }
        }

        /**
         * Test that against a real, Test file, we still do everything
         *  correctly
         */
        [Test]
        public void TestOnARealFile()
        {

            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("DateFormats.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            InternalWorkbook wb = workbook.Workbook;
            Assert.IsNotNull(wb);

            IRow row;
            ICell cell;
            NPOI.SS.UserModel.ICellStyle style;

            double aug_10_2007 = 39304.0;

            // Should have dates in 2nd column
            // All of them are the 10th of August
            // 2 US dates, 3 UK dates
            row = sheet.GetRow(0);
            cell = row.GetCell(1);
            style = cell.CellStyle;
            Assert.AreEqual(aug_10_2007, cell.NumericCellValue, 0.0001);
            Assert.AreEqual("d-mmm-yy", style.GetDataFormatString());
            Assert.IsTrue(DateUtil.IsInternalDateFormat(style.DataFormat));
            Assert.IsTrue(DateUtil.IsADateFormat(style.DataFormat, style.GetDataFormatString()));
            Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));

            row = sheet.GetRow(1);
            cell = row.GetCell(1);
            style = cell.CellStyle;
            Assert.AreEqual(aug_10_2007, cell.NumericCellValue, 0.0001);
            Assert.IsFalse(DateUtil.IsInternalDateFormat(cell.CellStyle.DataFormat));
            Assert.IsTrue(DateUtil.IsADateFormat(style.DataFormat, style.GetDataFormatString()));
            Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));

            row = sheet.GetRow(2);
            cell = row.GetCell(1);
            style = cell.CellStyle;
            Assert.AreEqual(aug_10_2007, cell.NumericCellValue, 0.0001);
            Assert.IsTrue(DateUtil.IsInternalDateFormat(cell.CellStyle.DataFormat));
            Assert.IsTrue(DateUtil.IsADateFormat(style.DataFormat, style.GetDataFormatString()));
            Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));

            row = sheet.GetRow(3);
            cell = row.GetCell(1);
            style = cell.CellStyle;
            Assert.AreEqual(aug_10_2007, cell.NumericCellValue, 0.0001);
            Assert.IsFalse(DateUtil.IsInternalDateFormat(cell.CellStyle.DataFormat));
            Assert.IsTrue(DateUtil.IsADateFormat(style.DataFormat, style.GetDataFormatString()));
            Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));

            row = sheet.GetRow(4);
            cell = row.GetCell(1);
            style = cell.CellStyle;
            Assert.AreEqual(aug_10_2007, cell.NumericCellValue, 0.0001);
            Assert.IsFalse(DateUtil.IsInternalDateFormat(cell.CellStyle.DataFormat));
            Assert.IsTrue(DateUtil.IsADateFormat(style.DataFormat, style.GetDataFormatString()));
            Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));
        }

        [Test]
        public void ExcelDateBorderCases()
        {
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");

            Assert.AreEqual(1.0, DateUtil.GetExcelDate(df.Parse("1900-01-01")), 0.00001);
            Assert.AreEqual(31.0, DateUtil.GetExcelDate(df.Parse("1900-01-31")), 0.00001);
            Assert.AreEqual(32.0, DateUtil.GetExcelDate(df.Parse("1900-02-01")), 0.00001);
            Assert.AreEqual(/* BAD_DATE! */ -1.0, DateUtil.GetExcelDate(df.Parse("1899-12-31")), 0.00001);
        }


        [Test]
        public void TestDateBug_2Excel()
        {
            Assert.AreEqual(59.0, DateUtil.GetExcelDate(new DateTime(1900, CALENDAR_FEBRUARY, 28), false), 1);
            Assert.AreEqual(61.0, DateUtil.GetExcelDate(new DateTime(1900, CALENDAR_MARCH, 1), false), 1);

            Assert.AreEqual(37315.00, DateUtil.GetExcelDate(new DateTime(2002, CALENDAR_FEBRUARY, 28), false), 1);
            Assert.AreEqual(37316.00, DateUtil.GetExcelDate(new DateTime(2002, CALENDAR_MARCH, 1), false), 1);
            Assert.AreEqual(37257.00, DateUtil.GetExcelDate(new DateTime(2002, CALENDAR_JANUARY, 1), false), 1);
            Assert.AreEqual(38074.00, DateUtil.GetExcelDate(new DateTime(2004, CALENDAR_MARCH, 28), false), 1);
        }
        [Test]
        public void TestDateBug_2Java()
        {
            Assert.AreEqual(new DateTime(1900, CALENDAR_FEBRUARY, 28), DateUtil.GetJavaDate(59.0, false));
            Assert.AreEqual(new DateTime(1900, CALENDAR_MARCH, 1), DateUtil.GetJavaDate(61.0, false));

            Assert.AreEqual(new DateTime(2002, CALENDAR_FEBRUARY, 28), DateUtil.GetJavaDate(37315.00, false));
            Assert.AreEqual(new DateTime(2002, CALENDAR_MARCH, 1), DateUtil.GetJavaDate(37316.00, false));
            Assert.AreEqual(new DateTime(2002, CALENDAR_JANUARY, 1), DateUtil.GetJavaDate(37257.00, false));
            Assert.AreEqual(new DateTime(2004, CALENDAR_MARCH, 28), DateUtil.GetJavaDate(38074.00, false));
        }
        [Test]
        public void TestDate1904()
        {
            Assert.AreEqual(new DateTime(1904, CALENDAR_JANUARY, 2), DateUtil.GetJavaDate(1.0, true));
            Assert.AreEqual(new DateTime(1904, CALENDAR_JANUARY, 1), DateUtil.GetJavaDate(0.0, true));
            Assert.AreEqual(0.0, DateUtil.GetExcelDate(new DateTime(1904, CALENDAR_JANUARY, 1), true), 0.00001);
            Assert.AreEqual(1.0, DateUtil.GetExcelDate(new DateTime(1904, CALENDAR_JANUARY, 2), true), 0.00001);

            Assert.AreEqual(new DateTime(1998, CALENDAR_JULY, 5), DateUtil.GetJavaDate(35981, false));
            Assert.AreEqual(new DateTime(1998, CALENDAR_JULY, 5), DateUtil.GetJavaDate(34519, true));

            Assert.AreEqual(35981.0, DateUtil.GetExcelDate(new DateTime(1998, CALENDAR_JULY, 5), false), 0.00001);
            Assert.AreEqual(34519.0, DateUtil.GetExcelDate(new DateTime(1998, CALENDAR_JULY, 5), true), 0.00001);
        }

        /**
         * Check if DateUtil.GetAbsoluteDay works as advertised.
         */
        [Test]
        public void TestAbsoluteDay()
        {
            // 1 Jan 1900 is 1 day after 31 Dec 1899
            DateTime calendar = new DateTime(1900, 1, 1);
            Assert.AreEqual(1, DateUtil.AbsoluteDay(calendar, false), "Checking absolute day (1 Jan 1900)");
            // 1 Jan 1901 is 366 days after 31 Dec 1899
            calendar = new DateTime(1901, 1, 1);
            Assert.AreEqual(366, DateUtil.AbsoluteDay(calendar, false), "Checking absolute day (1 Jan 1901)");
        }

        [Test]
        public void AbsoluteDayYearTooLow()
        {
            DateTime calendar = new DateTime(1899, 1, 1);
            try
            {
                HSSFDateUtil.AbsoluteDay(calendar, false);
                Assert.Fail("Should fail here");
            }
            catch (ArgumentException )
            {
                // expected here
            }

            try
            {
                calendar = new DateTime(1903, 1, 1);
                HSSFDateUtil.AbsoluteDay(calendar, true);
                Assert.Fail("Should fail here");
            }
            catch (ArgumentException )
            {
                // expected here
            }
        }
        [Test]
        public void TestConvertTime()
        {

            double delta = 1E-7; // a couple of digits more accuracy than strictly required
            Assert.AreEqual(0.5, DateUtil.ConvertTime("12:00"), delta);
            Assert.AreEqual(2.0 / 3, DateUtil.ConvertTime("16:00"), delta);
            Assert.AreEqual(0.0000116, DateUtil.ConvertTime("0:00:01"), delta);
            Assert.AreEqual(0.7330440, DateUtil.ConvertTime("17:35:35"), delta);
        }
        [Test]
        public void TestParseDate()
        {
            Assert.AreEqual(new DateTime(2008, 8, 3), DateUtil.ParseYYYYMMDDDate("2008/08/03"));
            Assert.AreEqual(new DateTime(1994, 5, 1), DateUtil.ParseYYYYMMDDDate("1994/05/01"));
        }

        /**
         * Ensure that date values *with* a fractional portion get the right time of day
         */
        [Test]
        public void TestConvertDateTime()
        {
            // Excel day 30000 is date 18-Feb-1982 
            // 0.7 corresponds to time 16:48:00
            DateTime actual = DateUtil.GetJavaDate(30000.7);
            DateTime expected = new DateTime(1982, 2, 18, 16, 48, 0);
            Assert.AreEqual(expected, actual);
        }
        /**
         * User reported a datetime issue in POI-2.5:
         *  Setting Cell's value to Jan 1, 1900 without a time doesn't return the same value Set to
         */
        [Test]
        public void TestBug19172()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            ICell cell = sheet.CreateRow(0).CreateCell(0);

            DateTime valueToTest = new DateTime(1900, 1, 1);

            cell.SetCellValue(valueToTest);

            DateTime returnedValue = (DateTime)cell.DateCellValue;

            Assert.AreEqual(valueToTest.TimeOfDay, returnedValue.TimeOfDay);
        }
        /**
         * DateUtil.isCellFormatted(Cell) should not true for a numeric cell 
         * that's formatted as ".0000"
         */
        [Test]
        public void TestBug54557()
        {
            string format = ".0000";
            bool isDateFormat = HSSFDateUtil.IsADateFormat(165, format);

            Assert.AreEqual(false, isDateFormat);
        }

        [Test]
        public void Bug56269()
        {
            double excelFraction = 41642.45833321759d;
            DateTime calNoRound = HSSFDateUtil.GetJavaCalendar(excelFraction, false);
            Assert.AreEqual(10, calNoRound.Hour);
            Assert.AreEqual(59, calNoRound.Minute);
            Assert.AreEqual(59, calNoRound.Second);
            DateTime calRound = HSSFDateUtil.GetJavaCalendar(excelFraction, false,(TimeZoneInfo)null, true);
            Assert.AreEqual(11, calRound.Hour);
            Assert.AreEqual(0, calRound.Minute);
            Assert.AreEqual(0, calRound.Second);
        }

    }
}

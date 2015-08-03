/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.UserModel
{
    using System.Globalization;
    using System;
    using System.Text.RegularExpressions;
    using System.Text;

    /// <summary>
    /// Contains methods for dealing with Excel dates.
    /// @author  Michael Harhen
    /// @author  Glen Stampoultzis (glens at apache.org)
    /// @author  Dan Sherman (dsherman at Isisph.com)
    /// @author  Hack Kampbjorn (hak at 2mba.dk)
    /// @author  Alex Jacoby (ajacoby at gmail.com)
    /// @author  Pavel Krupets (pkrupets at palmtreebusiness dot com)
    /// @author  Thies Wellpott
    /// </summary>
    public class DateUtil
    {
        public const int SECONDS_PER_MINUTE = 60;
        public const int MINUTES_PER_HOUR = 60;
        public const int HOURS_PER_DAY = 24;
        public const int SECONDS_PER_DAY = (HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE);

        private const int BAD_DATE = -1;   // used to specify that date Is invalid
        public const long DAY_MILLISECONDS = 24 * 60 * 60 * 1000;
        private static readonly char[] TIME_SEPARATOR_PATTERN = new char[] { ':' };

        /**
         * The following patterns are used in {@link #isADateFormat(int, String)}
         */
        private static Regex date_ptrn1 = new Regex("^\\[\\$\\-.*?\\]");
        private static Regex date_ptrn2 = new Regex("^\\[[a-zA-Z]+\\]");
        private static Regex date_ptrn3a = new Regex("[yYmMdDhHsS]");
        private static Regex date_ptrn3b = new Regex("^[\\[\\]yYmMdDhHsS\\-T/,. :\"\\\\]+0*[ampAMP/]*$");
        //  elapsed time patterns: [h],[m] and [s]
        //private static Regex date_ptrn4 = new Regex("^\\[([hH]+|[mM]+|[sS]+)\\]");
        private static Regex date_ptrn4 = new Regex("^\\[([hH]+|[mM]+|[sS]+)\\]$");


        /// <summary>
        /// Given a Calendar, return the number of days since 1899/12/31.
        /// </summary>
        /// <param name="cal">the date</param>
        /// <param name="use1904windowing">if set to <c>true</c> [use1904windowing].</param>
        /// <returns>number of days since 1899/12/31</returns>
        public static int AbsoluteDay(DateTime cal, bool use1904windowing)
        {
            int daynum = (cal - new DateTime(1899, 12, 31)).Days;
            if (cal > new DateTime(1900, 3, 1) && use1904windowing)
            {
                daynum++;
            }
            return daynum;
        }
        /// <summary>
        /// Given a Date, Converts it into a double representing its internal Excel representation,
        /// which Is the number of days since 1/1/1900. Fractional days represent hours, minutes, and seconds.
        /// </summary>
        /// <param name="date">Excel representation of Date (-1 if error - test for error by Checking for less than 0.1)</param>
        /// <returns>the Date</returns>
        public static double GetExcelDate(DateTime date)
        {
            return GetExcelDate(date, false);
        }
        /// <summary>
        /// Gets the excel date.
        /// </summary>
        /// <param name="year">The year.</param>
        /// <param name="month">The month.</param>
        /// <param name="day">The day.</param>
        /// <param name="hour">The hour.</param>
        /// <param name="minute">The minute.</param>
        /// <param name="second">The second.</param>
        /// <param name="use1904windowing">Should 1900 or 1904 date windowing be used?</param>
        /// <returns></returns>
        public static double GetExcelDate(int year, int month, int day, int hour, int minute, int second, bool use1904windowing)
        {
            if ((!use1904windowing && year < 1900)  //1900 date system must bigger than 1900
                || (use1904windowing && year < 1904))   //1904 date system must bigger than 1904
            {
                return BAD_DATE;
            }

            DateTime startdate;
            if (use1904windowing)
            {
                startdate = new DateTime(1904, 1, 1);
            }
            else
            {
                startdate = new DateTime(1900, 1, 1);
            }
            int nextyearmonth = 0;
            if (month > 12)
            {
                nextyearmonth = month - 12;
                month = 12;
            }
            int nextmonthday = 0;

            if ((month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12))
            {
                //big month
                if (day > 31)
                {
                    nextmonthday = day - 31;
                    day = 31;
                }
            }
            else if ((month == 4 || month == 6 || month == 9 || month == 11))
            {
                //small month
                if (day > 30)
                {
                    nextmonthday = day - 30;
                    day = 30;
                }
            }
            else if (DateTime.IsLeapYear(year))
            {
                //Feb. with leap year
                if (day > 29)
                {
                    nextmonthday = day - 29;
                    day = 29;
                }
            }
            else
            {
                //Feb without leap year
                if (day > 28)
                {
                    nextmonthday = day - 28;
                    day = 28;
                }
            }

            if (day <= 0)
            {
                nextmonthday = day - 1;
                day = 1;
            }

            DateTime date = new DateTime(year, month, day, hour, minute, second);
            date = date.AddMonths(nextyearmonth);
            date = date.AddDays(nextmonthday);
            double value = (date - startdate).TotalDays + 1;

            if (!use1904windowing && value >= 60)
            {
                value++;
            }
            else if (use1904windowing)
            {
                value--;
            }
            return value;
        }
        /// <summary>
        /// Given a Date, Converts it into a double representing its internal Excel representation,
        /// which Is the number of days since 1/1/1900. Fractional days represent hours, minutes, and seconds.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <param name="use1904windowing">Should 1900 or 1904 date windowing be used?</param>
        /// <returns>Excel representation of Date (-1 if error - test for error by Checking for less than 0.1)</returns>
        public static double GetExcelDate(DateTime date, bool use1904windowing)
        {
            if ((!use1904windowing && date.Year < 1900)  //1900 date system must bigger than 1900
                || (use1904windowing && date.Year < 1904))   //1904 date system must bigger than 1904
            {
                return BAD_DATE;
            }

            DateTime startdate;
            if (use1904windowing)
            {
                startdate = new DateTime(1904, 1, 1);
            }
            else
            {
                startdate = new DateTime(1900, 1, 1);
            }

            double value = (date - startdate).TotalDays + 1;

            if (!use1904windowing && value >= 60)
            {
                value++;
            }
            else if (use1904windowing)
            {
                value--;
            }
            return value;
        }

        /// <summary>
        ///  Given an Excel date with using 1900 date windowing, and converts it to a java.util.Date.
        ///  Excel Dates and Times are stored without any timezone 
        ///  information. If you know (through other means) that your file 
        ///  uses a different TimeZone to the system default, you can use
        ///  this version of the getJavaDate() method to handle it.
        /// </summary>
        /// <param name="date">The Excel date.</param>
        /// <returns>null if date is not a valid Excel date</returns>
        public static DateTime GetJavaDate(double date)
        {
            return GetJavaDate(date, false);
        }


        /**
         *  Given an Excel date with either 1900 or 1904 date windowing,
         *  Converts it to a Date.
         *
         *  NOTE: If the default <c>TimeZone</c> in Java uses Daylight
         *  Saving Time then the conversion back to an Excel date may not give
         *  the same value, that Is the comparison
         *  <CODE>excelDate == GetExcelDate(GetJavaDate(excelDate,false))</CODE>
         *  Is not always true. For example if default timezone Is
         *  <c>Europe/Copenhagen</c>, on 2004-03-28 the minute after
         *  01:59 CET Is 03:00 CEST, if the excel date represents a time between
         *  02:00 and 03:00 then it Is Converted to past 03:00 summer time
         *
         *  @param date  The Excel date.
         *  @param use1904windowing  true if date uses 1904 windowing,
         *   or false if using 1900 date windowing.
         *  @return Java representation of the date, or null if date Is not a valid Excel date
         *  @see TimeZone
         */
        public static DateTime GetJavaDate(double date, bool use1904windowing)
        {
            return GetJavaCalendar(date, use1904windowing, false);
        }
        /**
         *  Given an Excel date with either 1900 or 1904 date windowing,
         *  converts it to a java.util.Date.
         *  
         *  Excel Dates and Times are stored without any timezone 
         *  information. If you know (through other means) that your file 
         *  uses a different TimeZone to the system default, you can use
         *  this version of the getJavaDate() method to handle it.
         *   
         *  @param date  The Excel date.
         *  @param tz The TimeZone to evaluate the date in
         *  @param use1904windowing  true if date uses 1904 windowing,
         *   or false if using 1900 date windowing.
         *  @return Java representation of the date, or null if date is not a valid Excel date
         */
        public static DateTime getJavaDate(double date, bool use1904windowing, TimeZone tz)
        {
            return GetJavaCalendar(date, use1904windowing, false);
        }
        /**
         *  Given an Excel date with either 1900 or 1904 date windowing,
         *  converts it to a java.util.Date.
         *  
         *  Excel Dates and Times are stored without any timezone 
         *  information. If you know (through other means) that your file 
         *  uses a different TimeZone to the system default, you can use
         *  this version of the getJavaDate() method to handle it.
         *   
         *  @param date  The Excel date.
         *  @param tz The TimeZone to evaluate the date in
         *  @param use1904windowing  true if date uses 1904 windowing,
         *   or false if using 1900 date windowing.
         *  @param roundSeconds round to closest second
         *  @return Java representation of the date, or null if date is not a valid Excel date
         */
        public static DateTime GetJavaDate(double date, bool use1904windowing, TimeZone tz, bool roundSeconds)
        {
            return GetJavaCalendar(date, use1904windowing, roundSeconds);
        }

        public static void SetCalendar(ref DateTime calendar, int wholeDays,
            int millisecondsInDay, bool use1904windowing, bool roundSeconds)
        {
            int startYear = 1900;
            int dayAdjust = -1; // Excel thinks 2/29/1900 is a valid date, which it isn't
            if (use1904windowing)
            {
                startYear = 1904;
                dayAdjust = 1; // 1904 date windowing uses 1/2/1904 as the first day
            }
            else if (wholeDays < 61)
            {
                // Date is prior to 3/1/1900, so adjust because Excel thinks 2/29/1900 exists
                // If Excel date == 2/29/1900, will become 3/1/1900 in Java representation
                dayAdjust = 0;
            }
            DateTime dt = (new DateTime(startYear, 1, 1)).AddDays(wholeDays + dayAdjust - 1).AddMilliseconds(millisecondsInDay);
            if (roundSeconds)
            {
                dt = dt.AddMilliseconds(500);
                dt = dt.AddMilliseconds(-dt.Millisecond);
            }
            calendar = dt;

        }
        /**
         * Get EXCEL date as Java Calendar with given time zone.
         * @param date  The Excel date.
         * @param use1904windowing  true if date uses 1904 windowing,
         *  or false if using 1900 date windowing.
         * @param timeZone The TimeZone to evaluate the date in
         * @return Java representation of the date, or null if date is not a valid Excel date
         */
        public static DateTime GetJavaCalendar(double date, bool use1904windowing)
        {
            return GetJavaCalendar(date, use1904windowing, false);
        }
        /// <summary>
        /// Get EXCEL date as Java Calendar (with default time zone). This is like GetJavaDate(double, boolean) but returns a Calendar object.
        /// </summary>
        /// <param name="date">The Excel date.</param>
        /// <param name="use1904windowing">true if date uses 1904 windowing, or false if using 1900 date windowing.</param>
        /// <param name="roundSeconds"></param>
        /// <returns>null if date is not a valid Excel date</returns>
        public static DateTime GetJavaCalendar(double date, bool use1904windowing, bool roundSeconds)
        {
            if (!IsValidExcelDate(date))
            {
                throw new ArgumentException(string.Format(CultureInfo.CurrentCulture, "Invalid Excel date double value: {0}", date));
            }
            int wholeDays = (int)Math.Floor(date);
            int millisecondsInDay = (int)((date - wholeDays) * DAY_MILLISECONDS + 0.5);
            DateTime calendar;

            calendar = DateTime.Now;     // using default time-zone
            SetCalendar(ref calendar, wholeDays, millisecondsInDay, use1904windowing, roundSeconds);
            return calendar;
        }

        /// <summary>
        /// Converts a string of format "HH:MM" or "HH:MM:SS" to its (Excel) numeric equivalent
        /// </summary>
        /// <param name="timeStr">The time STR.</param>
        /// <returns> a double between 0 and 1 representing the fraction of the day</returns>
        public static double ConvertTime(String timeStr)
        {
            try
            {
                return ConvertTimeInternal(timeStr);
            }
            catch (FormatException e)
            {
                String msg = "Bad time format '" + timeStr
                    + "' expected 'HH:MM' or 'HH:MM:SS' - " + e.Message;
                throw new ArgumentException(msg);
            }
        }
        /// <summary>
        /// Converts the time internal.
        /// </summary>
        /// <param name="timeStr">The time STR.</param>
        /// <returns></returns>
        private static double ConvertTimeInternal(String timeStr)
        {
            int len = timeStr.Length;
            if (len < 4 || len > 8)
            {
                throw new FormatException("Bad length");
            }
            String[] parts = timeStr.Split(TIME_SEPARATOR_PATTERN);

            String secStr;
            switch (parts.Length)
            {
                case 2: secStr = "00"; break;
                case 3: secStr = parts[2]; break;
                default:
                    throw new FormatException("Expected 2 or 3 fields but got (" + parts.Length + ")");
            }
            String hourStr = parts[0];
            String minStr = parts[1];
            int hours = ParseInt(hourStr, "hour", HOURS_PER_DAY);
            int minutes = ParseInt(minStr, "minute", MINUTES_PER_HOUR);
            int seconds = ParseInt(secStr, "second", SECONDS_PER_MINUTE);

            double totalSeconds = seconds + (minutes + (hours) * 60) * 60;
            return totalSeconds / (SECONDS_PER_DAY);
        }

        // variables for performance optimization:
        // avoid re-checking DataUtil.isADateFormat(int, String) if a given format
        // string represents a date format if the same string is passed multiple times.
        // see https://issues.apache.org/bugzilla/show_bug.cgi?id=55611
        private static int lastFormatIndex = -1;
        private static String lastFormatString = null;
        private static bool cached = false;
        private static string syncIsADateFormat = "IsADateFormat";
        /// <summary>
        /// Given a format ID and its format String, will Check to see if the
        /// format represents a date format or not.
        /// Firstly, it will Check to see if the format ID corresponds to an
        /// internal excel date format (eg most US date formats)
        /// If not, it will Check to see if the format string only Contains
        /// date formatting Chars (ymd-/), which covers most
        /// non US date formats.
        /// </summary>
        /// <param name="formatIndex">The index of the format, eg from ExtendedFormatRecord.GetFormatIndex</param>
        /// <param name="formatString">The format string, eg from FormatRecord.GetFormatString</param>
        /// <returns>
        /// 	<c>true</c> if [is A date format] [the specified format index]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsADateFormat(int formatIndex, String formatString)
        {
            lock (syncIsADateFormat)
            {
                if (formatString != null && formatIndex == lastFormatIndex && formatString.Equals(lastFormatString))
                {
                    return cached;
                }
                // First up, Is this an internal date format?
                if (IsInternalDateFormat(formatIndex))
                {
                    lastFormatIndex = formatIndex;
                    lastFormatString = formatString;
                    cached = true;
                    return true;
                }

                // If we didn't get a real string, it can't be
                if (formatString == null || formatString.Length == 0)
                {
                    lastFormatIndex = formatIndex;
                    lastFormatString = formatString;
                    cached = false;
                    return false;
                }

                String fs = formatString;

                // If it end in ;@, that's some crazy dd/mm vs mm/dd
                //  switching stuff, which we can ignore
                fs = Regex.Replace(fs, ";@", "");
                StringBuilder sb = new StringBuilder(fs.Length);
                for (int i = 0; i < fs.Length; i++)
                {
                    char c = fs[i];
                    if (i < fs.Length - 1)
                    {
                        char nc = fs[i + 1];
                        if (c == '\\')
                        {
                            switch (nc)
                            {
                                case '-':
                                case ',':
                                case '.':
                                case ' ':
                                case '\\':
                                    // skip current '\' and continue to the next char
                                    continue;
                            }
                        }
                        else if (c == ';' && nc == '@')
                        {
                            i++;
                            // skip ";@" duplets
                            continue;
                        }
                    }
                    sb.Append(c);
                }
                fs = sb.ToString();


                // short-circuit if it indicates elapsed time: [h], [m] or [s]
                //if (Regex.IsMatch(fs, "^\\[([hH]+|[mM]+|[sS]+)\\]"))
                if (date_ptrn4.IsMatch(fs))
                {
                    lastFormatIndex = formatIndex;
                    lastFormatString = formatString;
                    cached = true;
                    return true;
                }

                // If it starts with [$-...], then could be a date, but
                //  who knows what that starting bit Is all about
                //fs = Regex.Replace(fs, "^\\[\\$\\-.*?\\]", "");
                fs = date_ptrn1.Replace(fs, "");

                // If it starts with something like [Black] or [Yellow],
                //  then it could be a date
                //fs = Regex.Replace(fs, "^\\[[a-zA-Z]+\\]", "");
                fs = date_ptrn2.Replace(fs, "");
                // You're allowed something like dd/mm/yy;[red]dd/mm/yy
                //  which would place dates before 1900/1904 in red
                // For now, only consider the first one
                if (fs.IndexOf(';') > 0 && fs.IndexOf(';') < fs.Length - 1)
                {
                    fs = fs.Substring(0, fs.IndexOf(';'));
                }
                // Ensure it has some date letters in it
                // (Avoids false positives on the rest of pattern 3)
                if (!date_ptrn3a.Match(fs).Success)
                //if (!Regex.Match(fs, "[yYmMdDhHsS]").Success)
                {
                    return false;
                }

                // If we get here, check it's only made up, in any case, of:
                //  y m d h s - \ / , . : [ ] T
                // optionally followed by AM/PM

                // Delete any string literals.
                fs = Regex.Replace(fs, @"""[^""\\]*(?:\\.[^""\\]*)*""", "");

                //if (Regex.IsMatch(fs, @"^[\[\]yYmMdDhHsS\-/,. :\""\\]+0*[ampAMP/]*$"))
                //{
                //    return true;
                //}

                //return false;

                bool result = date_ptrn3b.IsMatch(fs);
                lastFormatIndex = formatIndex;
                lastFormatString = formatString;
                cached = result;
                return result;
            }
        }
        /// <summary>
        /// Converts a string of format "YYYY/MM/DD" to its (Excel) numeric equivalent
        /// </summary>
        /// <param name="dateStr">The date STR.</param>
        /// <returns>a double representing the (integer) number of days since the start of the Excel epoch</returns>
        public static DateTime ParseYYYYMMDDDate(String dateStr)
        {
            try
            {
                return ParseYYYYMMDDDateInternal(dateStr);
            }
            catch (FormatException e)
            {
                String msg = "Bad time format " + dateStr
                    + " expected 'YYYY/MM/DD' - " + e.Message;
                throw new ArgumentException(msg);
            }
        }
        /// <summary>
        /// Parses the YYYYMMDD date internal.
        /// </summary>
        /// <param name="timeStr">The time string.</param>
        /// <returns></returns>
        private static DateTime ParseYYYYMMDDDateInternal(String timeStr)
        {
            if (timeStr.Length != 10)
            {
                throw new FormatException("Bad length");
            }

            String yearStr = timeStr.Substring(0, 4);
            String monthStr = timeStr.Substring(5, 2);
            String dayStr = timeStr.Substring(8, 2);
            int year = ParseInt(yearStr, "year", short.MinValue, short.MaxValue);
            int month = ParseInt(monthStr, "month", 1, 12);
            int day = ParseInt(dayStr, "day", 1, 31);

            DateTime cal = new DateTime(year, month, day, 0, 0, 0);
            return cal;
        }
        /// <summary>
        /// Parses the int.
        /// </summary>
        /// <param name="strVal">The string value.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="rangeMax">The range max.</param>
        /// <returns></returns>
        private static int ParseInt(String strVal, String fieldName, int rangeMax)
        {
            return ParseInt(strVal, fieldName, 0, rangeMax - 1);
        }
        /// <summary>
        /// Parses the int.
        /// </summary>
        /// <param name="strVal">The STR val.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="lowerLimit">The lower limit.</param>
        /// <param name="upperLimit">The upper limit.</param>
        /// <returns></returns>
        private static int ParseInt(String strVal, String fieldName, int lowerLimit, int upperLimit)
        {
            int result;
            try
            {
                result = int.Parse(strVal, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                throw new FormatException("Bad int format '" + strVal + "' for " + fieldName + " field");
            }
            if (result < lowerLimit || result > upperLimit)
            {
                throw new FormatException(fieldName + " value (" + result
                        + ") is outside the allowable range(0.." + upperLimit + ")");
            }
            return result;
        }
        /// <summary>
        /// Given a format ID this will Check whether the format represents an internal excel date format or not.
        /// </summary>
        /// <param name="format">The format.</param>
        public static bool IsInternalDateFormat(int format)
        {
            bool retval = false;

            switch (format)
            {
                // Internal Date Formats as described on page 427 in
                // Microsoft Excel Dev's Kit...
                case 0x0e:
                case 0x0f:
                case 0x10:
                case 0x11:
                case 0x12:
                case 0x13:
                case 0x14:
                case 0x15:
                case 0x16:
                case 0x2d:
                case 0x2e:
                case 0x2f:
                    retval = true;
                    break;

                default:
                    retval = false;
                    break;
            }
            return retval;
        }

        /// <summary>
        /// Check if a cell Contains a date
        /// Since dates are stored internally in Excel as double values
        /// we infer it Is a date if it Is formatted as such.
        /// </summary>
        /// <param name="cell">The cell.</param>
        public static bool IsCellDateFormatted(ICell cell)
        {
            if (cell == null) return false;
            bool bDate = false;

            double d = cell.NumericCellValue;
            if (DateUtil.IsValidExcelDate(d))
            {
                ICellStyle style = cell.CellStyle;
                if (style == null)
                    return false;
                int i = style.DataFormat;
                String f = style.GetDataFormatString();
                bDate = IsADateFormat(i, f);
            }
            return bDate;
        }
        /// <summary>
        /// Check if a cell contains a date, Checking only for internal excel date formats.
        /// As Excel stores a great many of its dates in "non-internal" date formats, you will not normally want to use this method.
        /// </summary>
        /// <param name="cell">The cell.</param>
        public static bool IsCellInternalDateFormatted(ICell cell)
        {
            if (cell == null) return false;
            bool bDate = false;

            double d = cell.NumericCellValue;
            if (DateUtil.IsValidExcelDate(d))
            {
                ICellStyle style = cell.CellStyle;
                int i = style.DataFormat;
                bDate = IsInternalDateFormat(i);
            }
            return bDate;
        }


        /// <summary>
        /// Given a double, Checks if it Is a valid Excel date.
        /// </summary>
        /// <param name="value">the double value.</param>
        /// <returns>
        /// 	<c>true</c> if [is valid excel date] [the specified value]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsValidExcelDate(double value)
        {
            //return true;
            return value > -Double.Epsilon;
        }

    }
}
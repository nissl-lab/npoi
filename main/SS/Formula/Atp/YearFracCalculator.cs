/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.Formula.Atp
{
    using System;
    using NPOI.SS.Formula.Eval;


    /// <summary>
    /// Internal calculation methods for Excel 'Analysis ToolPak' function YEARFRAC()
    /// Algorithm inspired by www.dwheeler.com/yearfrac
    /// @author Josh Micich
    /// </summary>
    /// <remarks>
    /// Date Count convention 
    /// http://en.wikipedia.org/wiki/Day_count_convention
    /// </remarks>
    /// <remarks>
    /// Office Online Help on YEARFRAC
    /// http://office.microsoft.com/en-us/excel/HP052093441033.aspx
    /// </remarks>

    public class YearFracCalculator
    {
        /** use UTC time-zone to avoid daylight savings issues */
        //private static readonly TimeZone UTC_TIME_ZONE = TimeZone.GetTimeZone("UTC");
        private const int MS_PER_HOUR = 60 * 60 * 1000;
        private const int MS_PER_DAY = 24 * MS_PER_HOUR;
        private const int DAYS_PER_NORMAL_YEAR = 365;
        private const int DAYS_PER_LEAP_YEAR = DAYS_PER_NORMAL_YEAR + 1;

        /** the length of normal long months i.e. 31 */
        private const int LONG_MONTH_LEN = 31;
        /** the length of normal short months i.e. 30 */
        private const int SHORT_MONTH_LEN = 30;
        private const int SHORT_FEB_LEN = 28;
        private const int LONG_FEB_LEN = SHORT_FEB_LEN + 1;


        /// <summary>
        /// Calculates YEARFRAC()
        /// </summary>
        /// <param name="pStartDateVal">The start date.</param>
        /// <param name="pEndDateVal">The end date.</param>
        /// <param name="basis">The basis value.</param>
        /// <returns></returns>
        public static double Calculate(double pStartDateVal, double pEndDateVal, int basis)
        {

            if (basis < 0 || basis >= 5)
            {
                // if basis is invalid the result is #NUM!
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }

            // common logic for all bases

            // truncate day values
            int startDateVal = (int)Math.Floor(pStartDateVal);
            int endDateVal = (int)Math.Floor(pEndDateVal);
            if (startDateVal == endDateVal)
            {
                // when dates are equal, result is zero 
                return 0;
            }
            // swap start and end if out of order
            if (startDateVal > endDateVal)
            {
                int temp = startDateVal;
                startDateVal = endDateVal;
                endDateVal = temp;
            }

            switch (basis)
            {
                case 0: return Basis0(startDateVal, endDateVal);
                case 1: return Basis1(startDateVal, endDateVal);
                case 2: return Basis2(startDateVal, endDateVal);
                case 3: return Basis3(startDateVal, endDateVal);
                case 4: return Basis4(startDateVal, endDateVal);
            }
            throw new InvalidOperationException("cannot happen");
        }


        /// <summary>
        /// Basis 0, 30/360 date convention 
        /// </summary>
        /// <param name="startDateVal">The start date value assumed to be less than or equal to endDateVal.</param>
        /// <param name="endDateVal">The end date value assumed to be greater than or equal to startDateVal.</param>
        /// <returns></returns>
        public static double Basis0(int startDateVal, int endDateVal)
        {
            SimpleDate startDate = CreateDate(startDateVal);
            SimpleDate endDate = CreateDate(endDateVal);
            int date1day = startDate.day;
            int date2day = endDate.day;

            // basis zero has funny adjustments to the day-of-month fields when at end-of-month 
            if (date1day == LONG_MONTH_LEN && date2day == LONG_MONTH_LEN)
            {
                date1day = SHORT_MONTH_LEN;
                date2day = SHORT_MONTH_LEN;
            }
            else if (date1day == LONG_MONTH_LEN)
            {
                date1day = SHORT_MONTH_LEN;
            }
            else if (date1day == SHORT_MONTH_LEN && date2day == LONG_MONTH_LEN)
            {
                date2day = SHORT_MONTH_LEN;
                // Note: If date2day==31, it STAYS 31 if date1day < 30.
                // Special fixes for February:
            }
            else if (startDate.month == 2 && IsLastDayOfMonth(startDate))
            {
                // Note - these assignments deliberately set Feb 30 date.
                date1day = SHORT_MONTH_LEN;
                if (endDate.month == 2 && IsLastDayOfMonth(endDate))
                {
                    // only adjusted when first date is last day in Feb
                    date2day = SHORT_MONTH_LEN;
                }
            }
            return CalculateAdjusted(startDate, endDate, date1day, date2day);
        }
        /// <summary>
        /// Basis 1, Actual/Actual date convention 
        /// </summary>
        /// <param name="startDateVal">The start date value assumed to be less than or equal to endDateVal.</param>
        /// <param name="endDateVal">The end date value assumed to be greater than or equal to startDateVal.</param>
        /// <returns></returns>
        public static double Basis1(int startDateVal, int endDateVal)
        {
            SimpleDate startDate = CreateDate(startDateVal);
            SimpleDate endDate = CreateDate(endDateVal);
            double yearLength;
            if (IsGreaterThanOneYear(startDate, endDate))
            {
                yearLength = AverageYearLength(startDate.year, endDate.year);
            }
            else if (ShouldCountFeb29(startDate, endDate))
            {
                yearLength = DAYS_PER_LEAP_YEAR;
            }
            else
            {
                yearLength = DAYS_PER_NORMAL_YEAR;
            }
            return DateDiff(startDate.ticks, endDate.ticks) / yearLength;
        }

        /// <summary>
        /// Basis 2, Actual/360 date convention 
        /// </summary>
        /// <param name="startDateVal">The start date value assumed to be less than or equal to endDateVal.</param>
        /// <param name="endDateVal">The end date value assumed to be greater than or equal to startDateVal.</param>
        /// <returns></returns>
        public static double Basis2(int startDateVal, int endDateVal)
        {
            return (endDateVal - startDateVal) / 360.0;
        }
        /// <summary>
        /// Basis 3, Actual/365 date convention 
        /// </summary>
        /// <param name="startDateVal">The start date value assumed to be less than or equal to endDateVal.</param>
        /// <param name="endDateVal">The end date value assumed to be greater than or equal to startDateVal.</param>
        /// <returns></returns>
        public static double Basis3(double startDateVal, double endDateVal)
        {
            return (endDateVal - startDateVal) / 365.0;
        }
        /// <summary>
        /// Basis 4, European 30/360 date convention 
        /// </summary>
        /// <param name="startDateVal">The start date value assumed to be less than or equal to endDateVal.</param>
        /// <param name="endDateVal">The end date value assumed to be greater than or equal to startDateVal.</param>
        /// <returns></returns>
        public static double Basis4(int startDateVal, int endDateVal)
        {
            SimpleDate startDate = CreateDate(startDateVal);
            SimpleDate endDate = CreateDate(endDateVal);
            int date1day = startDate.day;
            int date2day = endDate.day;


            // basis four has funny adjustments to the day-of-month fields when at end-of-month 
            if (date1day == LONG_MONTH_LEN)
            {
                date1day = SHORT_MONTH_LEN;
            }
            if (date2day == LONG_MONTH_LEN)
            {
                date2day = SHORT_MONTH_LEN;
            }
            // Note - no adjustments for end of Feb
            return CalculateAdjusted(startDate, endDate, date1day, date2day);
        }


        /// <summary>
        /// Calculates the adjusted.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        /// <param name="endDate">The end date.</param>
        /// <param name="date1day">The date1day.</param>
        /// <param name="date2day">The date2day.</param>
        /// <returns></returns>
        private static double CalculateAdjusted(SimpleDate startDate, SimpleDate endDate, int date1day,
                int date2day)
        {
            double dayCount
                = (endDate.year - startDate.year) * 360
                + (endDate.month - startDate.month) * SHORT_MONTH_LEN
                + (date2day - date1day) * 1;
            return dayCount / 360;
        }

        /// <summary>
        /// Determines whether [is last day of month] [the specified date].
        /// </summary>
        /// <param name="date">The date.</param>
        /// <returns>
        /// 	<c>true</c> if [is last day of month] [the specified date]; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsLastDayOfMonth(SimpleDate date)
        {
            if (date.day < SHORT_FEB_LEN)
            {
                return false;
            }
            return date.day == GetLastDayOfMonth(date);
        }

        /// <summary>
        /// Gets the last day of month.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <returns></returns>
        private static int GetLastDayOfMonth(SimpleDate date)
        {
            switch (date.month)
            {
                case 1:
                case 3:
                case 5:
                case 7:
                case 8:
                case 10:
                case 12:
                    return LONG_MONTH_LEN;
                case 4:
                case 6:
                case 9:
                case 11:
                    return SHORT_MONTH_LEN;
            }
            if (IsLeapYear(date.year))
            {
                return LONG_FEB_LEN;
            }
            return SHORT_FEB_LEN;
        }

        /// <summary>
        /// Assumes dates are no more than 1 year apart.
        /// </summary>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <returns><c>true</c>
        ///  if dates both within a leap year, or span a period including Feb 29</returns>
        private static bool ShouldCountFeb29(SimpleDate start, SimpleDate end)
        {
            bool startIsLeapYear = IsLeapYear(start.year);
            if (startIsLeapYear && start.year == end.year)
            {
                // note - dates may not actually span Feb-29, but it gets counted anyway in this case
                return true;
            }

            bool endIsLeapYear = IsLeapYear(end.year);
            if (!startIsLeapYear && !endIsLeapYear)
            {
                return false;
            }
            if (startIsLeapYear)
            {
                switch (start.month)
                {
                    case SimpleDate.JANUARY:
                    case SimpleDate.FEBRUARY:
                        return true;
                }
                return false;
            }
            if (endIsLeapYear)
            {
                switch (end.month)
                {
                    case SimpleDate.JANUARY:
                        return false;
                    case SimpleDate.FEBRUARY:
                        break;
                    default:
                        return true;
                }
                return end.day == LONG_FEB_LEN;
            }
            return false;
        }

        /// <summary>
        /// return the whole number of days between the two time-stamps.  Both time-stamps are
        /// assumed to represent 12:00 midnight on the respective day.
        /// </summary>
        /// <param name="startDateTicks">The start date ticks.</param>
        /// <param name="endDateTicks">The end date ticks.</param>
        /// <returns></returns>
        private static double DateDiff(long startDateTicks, long endDateTicks)
        {
            return new TimeSpan(endDateTicks - startDateTicks).TotalDays;
        }

        /// <summary>
        /// Averages the length of the year.
        /// </summary>
        /// <param name="startYear">The start year.</param>
        /// <param name="endYear">The end year.</param>
        /// <returns></returns>
        private static double AverageYearLength(int startYear, int endYear)
        {
            int dayCount = 0;
            for (int i = startYear; i <= endYear; i++)
            {
                dayCount += DAYS_PER_NORMAL_YEAR;
                if (IsLeapYear(i))
                {
                    dayCount++;
                }
            }
            double numberOfYears = endYear - startYear + 1;
            return dayCount / numberOfYears;
        }

        /// <summary>
        /// determine Leap Year
        /// </summary>
        /// <param name="i">the year</param>
        /// <returns></returns>
        private static bool IsLeapYear(int i)
        {
            // leap years are always divisible by 4
            if (i % 4 != 0)
            {
                return false;
            }
            // each 4th century is a leap year
            if (i % 400 == 0)
            {
                return true;
            }
            // all other centuries are *not* leap years
            if (i % 100 == 0)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Determines whether [is greater than one year] [the specified start].
        /// </summary>
        /// <param name="start">The start date.</param>
        /// <param name="end">The end date.</param>
        /// <returns>
        /// 	<c>true</c> if [is greater than one year] [the specified start]; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsGreaterThanOneYear(SimpleDate start, SimpleDate end)
        {
            if (start.year == end.year)
            {
                return false;
            }
            if (start.year + 1 != end.year)
            {
                return true;
            }

            if (start.month > end.month)
            {
                return false;
            }
            if (start.month < end.month)
            {
                return true;
            }

            return start.day < end.day;
        }

        /// <summary>
        /// Creates the date.
        /// </summary>
        /// <param name="dayCount">The day count.</param>
        /// <returns></returns>
        private static SimpleDate CreateDate(int dayCount)
        {
            DateTime dt = DateTime.Now;
            NPOI.SS.UserModel.DateUtil.SetCalendar(ref dt, dayCount, 0, false, false);
            return new SimpleDate(dt);
        }

        /// <summary>
        /// Simple Date Wrapper
        /// </summary>
        private class SimpleDate
        {

            public const int JANUARY = 1;
            public const int FEBRUARY = 2;

            public int year;
            /** 1-based month */
            public int month;
            /** day of month */
            public int day;
            /** milliseconds since 1970 */
            public long ticks;

            public SimpleDate(DateTime date)
            {
                year = date.Year;
                month = date.Month;
                day = date.Day;
                ticks = date.Ticks;
            }
        }
    }
}

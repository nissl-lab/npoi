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
using System;
using NPOI.SS.UserModel;
namespace NPOI.SS.Formula.Atp
{
    /**
     * A calculator for workdays, considering dates as excel representations.
     * 
     * @author jfaenomoto@gmail.com
     */
    public class WorkdayCalculator
    {

        public static WorkdayCalculator instance = new WorkdayCalculator();

        /**
         * Constructor.
         */
        private WorkdayCalculator()
        {
            // enforcing singleton
        }

        /**
         * Calculate how many workdays are there between a start and an end date, as excel representations, considering a range of holidays.
         * 
         * @param start start date.
         * @param end end date.
         * @param holidays an array of holidays.
         * @return number of workdays between start and end dates, including both dates.
         */
        public int CalculateWorkdays(double start, double end, double[] holidays)
        {
            int saturdaysPast = this.PastDaysOfWeek(start, end, DayOfWeek.Saturday);
            int sundaysPast = this.PastDaysOfWeek(start, end, DayOfWeek.Sunday);
            int nonWeekendHolidays = this.CalculateNonWeekendHolidays(start, end, holidays);
            return (int)(end - start + 1) - saturdaysPast - sundaysPast - nonWeekendHolidays;
        }

        /**
         * Calculate the workday past x workdays from a starting date, considering a range of holidays.
         * 
         * @param start start date.
         * @param workdays number of workdays to be past from starting date.
         * @param holidays an array of holidays.
         * @return date past x workdays.
         */
        public DateTime CalculateWorkdays(double start, int workdays, double[] holidays)
        {
            DateTime startDate = DateUtil.GetJavaDate(start);
            //Calendar endDate = Calendar.getInstance();
            //endDate.setTime(startDate);
            int direction = workdays < 0 ? -1 : 1;
            DateTime endDate = startDate;
            double excelEndDate = DateUtil.GetExcelDate(endDate);
            while (workdays != 0)
            {
                endDate = endDate.AddDays(direction);
                excelEndDate += direction;
                if (endDate.DayOfWeek!= DayOfWeek.Saturday
                        && endDate.DayOfWeek != DayOfWeek.Sunday
                        && !IsHoliday(excelEndDate, holidays))
                {
                    workdays -= direction;
                }
            }
            //DateTime endDate = startDate.AddDays(workdays);
            //int skippedDays = 0;
            //do
            //{
            //    double end = DateUtil.GetExcelDate(endDate);
            //    int saturdaysPast = this.PastDaysOfWeek(start, end, DayOfWeek.Saturday);
            //    int sundaysPast = this.PastDaysOfWeek(start, end, DayOfWeek.Sunday);
            //    int nonWeekendHolidays = this.CalculateNonWeekendHolidays(start, end, holidays);
            //    skippedDays = saturdaysPast + sundaysPast + nonWeekendHolidays;
            //    endDate = endDate.AddDays(skippedDays);
            //    start = end + IsNonWorkday(end, holidays);
            //} while (skippedDays != 0);
            return endDate;
        }

        /**
         * Calculates how many days of week past between a start and an end date.
         * 
         * @param start start date.
         * @param end end date.
         * @param dayOfWeek a day of week as represented by {@link Calendar} constants.
         * @return how many days of week past in this interval.
         */
        public int PastDaysOfWeek(double start, double end, DayOfWeek dayOfWeek)
        {
            int pastDaysOfWeek = 0;
            int startDay = (int)Math.Floor(start < end ? start : end);
            int endDay = (int)Math.Floor(end > start ? end : start);
            for (; startDay <= endDay; startDay++)
            {
                DateTime today = DateUtil.GetJavaDate(startDay);
                if (today.DayOfWeek == dayOfWeek)
                {
                    pastDaysOfWeek++;
                }
            }
            return start < end ? pastDaysOfWeek : -pastDaysOfWeek;
        }

        /**
         * Calculates how many holidays in a list are workdays, considering an interval of dates.
         * 
         * @param start start date.
         * @param end end date.
         * @param holidays an array of holidays.
         * @return number of holidays that occur in workdays, between start and end dates.
         */

        private int CalculateNonWeekendHolidays(double start, double end, double[] holidays)
        {
            int nonWeekendHolidays = 0;
            double startDay = start < end ? start : end;
            double endDay = end > start ? end : start;
            for (int i = 0; i < holidays.Length; i++)
            {
                if (IsInARange(startDay, endDay, holidays[i]))
                {
                    if (!IsWeekend(holidays[i]))
                    {
                        nonWeekendHolidays++;
                    }
                }
            }
            return start < end ? nonWeekendHolidays : -nonWeekendHolidays;
        }

        /**
         * @param aDate a given date.
         * @return <code>true</code> if date is weekend, <code>false</code> otherwise.
         */

        private bool IsWeekend(double aDate)
        {
            DateTime date = DateUtil.GetJavaDate(aDate);
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }

        /**
         * @param aDate a given date.
         * @param holidays an array of holidays.
         * @return <code>true</code> if date is a holiday, <code>false</code> otherwise.
         */

        private bool IsHoliday(double aDate, double[] holidays)
        {
            for (int i = 0; i < holidays.Length; i++)
            {
                if (Math.Round(holidays[i]) == Math.Round(aDate))
                {
                    return true;
                }
            }
            return false;
        }

        /**
         * @param aDate a given date.
         * @param holidays an array of holidays.
         * @return <code>1</code> is not a workday, <code>0</code> otherwise.
         */

        private int IsNonWorkday(double aDate, double[] holidays)
        {
            return IsWeekend(aDate) || IsHoliday(aDate, holidays) ? 1 : 0;
        }

        /**
         * @param start start date.
         * @param end end date.
         * @param aDate a date to be analyzed.
         * @return <code>true</code> if aDate is between start and end dates, <code>false</code> otherwise.
         */

        private bool IsInARange(double start, double end, double aDate)
        {
            return aDate >= start && aDate <= end;
        }
    }
}
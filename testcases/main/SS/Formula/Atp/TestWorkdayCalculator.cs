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
using System.Diagnostics.CodeAnalysis;
using NPOI.SS.Formula.Atp;
using NPOI.SS.UserModel;
using NUnit.Framework;

namespace TestCases.SS.Formula.Atp
{
    [TestFixture]
    public class TestWorkdayCalculator
    {
        private const int December = 12;

        [Test]
        public void TestCalculateWorkdaysShouldReturnJustWeekdaysWhenNoWeekend()
        {
            double A_MONDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 12));
            double A_FRIDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 16));
            Assert.AreEqual(5, WorkdayCalculator.instance.CalculateWorkdays(A_MONDAY, A_FRIDAY, new double[0]));
        }
        [Test]
        public void TestCalculateWorkdaysShouldReturnAllDaysButNoSaturdays()
        {
            double A_WEDNESDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 14));
            double A_SATURDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 18));
            Assert.AreEqual(3, WorkdayCalculator.instance.CalculateWorkdays(A_WEDNESDAY, A_SATURDAY, new double[0]));
        }
        [Test]
        public void TestCalculateWorkdaysShouldReturnAllDaysButNoSundays()
        {
            double A_SUNDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 11));
            double A_THURSDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 15));
            Assert.AreEqual(4, WorkdayCalculator.instance.CalculateWorkdays(A_SUNDAY, A_THURSDAY, new double[0]));
        }
        [Test]
        public void TestCalculateWorkdaysShouldReturnAllDaysButNoHolidays()
        {
            double A_MONDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 12));
            double A_FRIDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 16));
            double A_WEDNESDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 14));
            Assert.AreEqual(4, WorkdayCalculator.instance.CalculateWorkdays(A_MONDAY, A_FRIDAY, new double[] { A_WEDNESDAY }));
        }
        [Test]
        public void TestCalculateWorkdaysShouldIgnoreWeekendHolidays()
        {
            double A_FRIDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 16));
            double A_SATURDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 17));
            double A_SUNDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 18));
            double A_WEDNESDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 21));
            Assert.AreEqual(4, WorkdayCalculator.instance.CalculateWorkdays(A_FRIDAY, A_WEDNESDAY, new double[] { A_SATURDAY, A_SUNDAY }));
        }

        [Test]
        public void TestCalculateWorkdaysOnSameDayShouldReturn1ForWeekdays()
        {
            double A_MONDAY = DateUtil.GetExcelDate(new DateTime(2017, 1, 2));
            Assert.AreEqual(1, WorkdayCalculator.instance.CalculateWorkdays(A_MONDAY, A_MONDAY, new double[0]));
        }

        [Test]
        public void TestCalculateWorkdaysOnSameDayShouldReturn0ForHolidays()
        {
            double A_MONDAY = DateUtil.GetExcelDate(new DateTime(2017, 1, 2));
            Assert.AreEqual(0, WorkdayCalculator.instance.CalculateWorkdays(A_MONDAY, A_MONDAY, new double[] { A_MONDAY }));
        }

        [Test]
        public void TestCalculateWorkdaysOnSameDayShouldReturn0ForWeekends()
        {
            double A_SUNDAY = DateUtil.GetExcelDate(new DateTime(2017, 1, 1));
            Assert.AreEqual(0, WorkdayCalculator.instance.CalculateWorkdays(A_SUNDAY, A_SUNDAY, new double[0]));
        }

        [Test]
        public void TestCalculateWorkdaysNumberOfDays()
        {
            double start = 41553.0;
            int days = 1;
            Assert.AreEqual(new DateTime(2013, 10, 7), WorkdayCalculator.instance.CalculateWorkdays(start, days, new double[0]));
        }


        [Test]
        public void TestPastDaysOfWeekShouldReturn0Past0Saturdays()
        {
            double A_WEDNESDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 7));
            double A_FRIDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 9));
            Assert.AreEqual(0, WorkdayCalculator.instance.PastDaysOfWeek(A_WEDNESDAY, A_FRIDAY, DayOfWeek.Saturday));
        }
        [Test]
        public void TestPastDaysOfWeekShouldReturn1Past1Saturdays()
        {
            double A_WEDNESDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 7));
            double A_SUNDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 11));
            Assert.AreEqual(1, WorkdayCalculator.instance.PastDaysOfWeek(A_WEDNESDAY, A_SUNDAY,  DayOfWeek.Saturday));
        }
        [Test]
        public void TestPastDaysOfWeekShouldReturn2Past2Saturdays()
        {
            double A_THURSDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 8));
            double A_MONDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 19));
            Assert.AreEqual(2, WorkdayCalculator.instance.PastDaysOfWeek(A_THURSDAY, A_MONDAY,  DayOfWeek.Saturday));
        }
        [Test]
        public void TestPastDaysOfWeekShouldReturn1BeginningFromASaturday()
        {
            double A_SATURDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 10));
            double A_SUNDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 11));
            Assert.AreEqual(1, WorkdayCalculator.instance.PastDaysOfWeek(A_SATURDAY, A_SUNDAY,  DayOfWeek.Saturday));
        }
        [Test]
        public void TestPastDaysOfWeekShouldReturn1EndingAtASaturday()
        {
            double A_THURSDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 8));
            double A_SATURDAY = DateUtil.GetExcelDate(new DateTime(2011, December, 10));
            Assert.AreEqual(1, WorkdayCalculator.instance.PastDaysOfWeek(A_THURSDAY, A_SATURDAY,  DayOfWeek.Saturday));
        }

        [Test]
        public void TestCalculateNonWeekendHolidays()
        {
            double start = DateUtil.GetExcelDate(new DateTime(2016, 12, 24));
            double end = DateUtil.GetExcelDate(new DateTime(2016, 12, 31));
            double holiday1 = DateUtil.GetExcelDate(new DateTime(2016, 12, 25));
            double holiday2 = DateUtil.GetExcelDate(new DateTime(2016, 12, 26));
            int count = WorkdayCalculator.instance.CalculateNonWeekendHolidays(start, end, new double[] { holiday1, holiday2 });
            Assert.AreEqual(1, count,
                "Expected 1 non-weekend-holiday for " + start + " to " + end + " and " + holiday1 + " and " + holiday2);
        }

        [Test]
        public void TestCalculateNonWeekendHolidaysOneDay()
        {
            double start = DateUtil.GetExcelDate(new DateTime(2016, 12, 26));
            double end = DateUtil.GetExcelDate(new DateTime(2016, 12, 26));
            double holiday1 = DateUtil.GetExcelDate(new DateTime(2016, 12, 25));
            double holiday2 = DateUtil.GetExcelDate(new DateTime(2016, 12, 26));
            int count = WorkdayCalculator.instance.CalculateNonWeekendHolidays(start, end, new double[] { holiday1, holiday2 });
            Assert.AreEqual(1, count,
                "Expected 1 non-weekend-holiday for " + start + " to " + end + " and " + holiday1 + " and " + holiday2);
        }

        [Test]
        public void TestIsNonWorkday()
        {
            double weekend = DateUtil.GetExcelDate(new DateTime(2016, 12, 25));
            double holiday = DateUtil.GetExcelDate(new DateTime(2016, 12, 26));
            double workday = DateUtil.GetExcelDate(new DateTime(2016, 12, 27));
#pragma warning disable CS0618 // 类型或成员已过时
            Assert.AreEqual(1, WorkdayCalculator.instance.IsNonWorkday(weekend, new double[] { holiday }));
            Assert.AreEqual(1, WorkdayCalculator.instance.IsNonWorkday(holiday, new double[] { holiday }));
            Assert.AreEqual(0, WorkdayCalculator.instance.IsNonWorkday(workday, new double[] { holiday }));
#pragma warning restore CS0618 // 类型或成员已过时

        }
    }

}
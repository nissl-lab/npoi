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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.SS.UserModel
{
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NUnit.Framework;
    using System.Globalization;

    public class TestDateUtil
    {

        private void EnsureDateException()
        {

        }

        [Test]
        public void GetJavaDate_InvalidValue()
        {
            double dateValue = -1;
            TimeZoneInfo tz = LocaleUtil.GetUserTimeZoneInfo();
            bool use1904windowing = false;
            bool roundSeconds = false;

            try
            {
                DateUtil.GetJavaDate(dateValue);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            try
            {
                DateUtil.GetJavaDate(dateValue, tz);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            try
            {
                DateUtil.GetJavaDate(dateValue, use1904windowing);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            try
            {
                DateUtil.GetJavaDate(dateValue, use1904windowing, tz);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            try
            {
                DateUtil.GetJavaDate(dateValue, use1904windowing, tz, roundSeconds);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            //Assert.AreEqual(null, DateUtil.GetJavaDate(dateValue));
            //Assert.AreEqual(null, DateUtil.GetJavaDate(dateValue, tz));
            //Assert.AreEqual(null, DateUtil.GetJavaDate(dateValue, use1904windowing));
            //Assert.AreEqual(null, DateUtil.GetJavaDate(dateValue, use1904windowing, tz));
            //Assert.AreEqual(null, DateUtil.GetJavaDate(dateValue, use1904windowing, tz, roundSeconds));
        }

        [Test]
        public void GetJavaDate_ValidValue()
        {
            double dateValue = 0;
            TimeZoneInfo tz = LocaleUtil.GetUserTimeZoneInfo();
            bool use1904windowing = false;
            bool roundSeconds = false;

            DateTime calendar = LocaleUtil.GetLocaleCalendar(1900, 1, 0);
            //Date date = calendar.GetTime();
            DateTime date = calendar;

            Assert.AreEqual(date, DateUtil.GetJavaDate(dateValue));
            Assert.AreEqual(date, DateUtil.GetJavaDate(dateValue, tz));
            Assert.AreEqual(date, DateUtil.GetJavaDate(dateValue, use1904windowing));
            Assert.AreEqual(date, DateUtil.GetJavaDate(dateValue, use1904windowing, tz));
            Assert.AreEqual(date, DateUtil.GetJavaDate(dateValue, use1904windowing, tz, roundSeconds));
        }

        [Test]
        public void GetJavaCalendar_InvalidValue()
        {
            double dateValue = -1;
            TimeZoneInfo tz = LocaleUtil.GetUserTimeZoneInfo();
            bool use1904windowing = false;
            bool roundSeconds = false;

            try
            {
                DateUtil.GetJavaDate(dateValue);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            try
            {
                DateUtil.GetJavaDate(dateValue, use1904windowing);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            try
            {
                DateUtil.GetJavaDate(dateValue, use1904windowing, tz);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            try
            {
                DateUtil.GetJavaDate(dateValue, use1904windowing, tz, roundSeconds);
                Assert.Fail("invalid datetime double value -1");
            }
            catch (ArgumentException) { }
            //Assert.AreEqual(null, DateUtil.GetJavaCalendar(dateValue));
            //Assert.AreEqual(null, DateUtil.GetJavaCalendar(dateValue, use1904windowing));
            //Assert.AreEqual(null, DateUtil.GetJavaCalendar(dateValue, use1904windowing, tz));
            //Assert.AreEqual(null, DateUtil.GetJavaCalendar(dateValue, use1904windowing, tz, roundSeconds));
        }

        [Test]
        public void GetJavaCalendar_ValidValue()
        {
            double dateValue = 0;
            TimeZoneInfo tz = LocaleUtil.GetUserTimeZoneInfo();
            bool use1904windowing = false;
            bool roundSeconds = false;

            DateTime expCal = LocaleUtil.GetLocaleCalendar(1900, 1, 0);

            DateTime[] actCal = {
            DateUtil.GetJavaCalendar(dateValue),
            DateUtil.GetJavaCalendar(dateValue, use1904windowing),
            DateUtil.GetJavaCalendar(dateValue, use1904windowing, tz),
            DateUtil.GetJavaCalendar(dateValue, use1904windowing, tz, roundSeconds)
        };
            Assert.AreEqual(expCal, actCal[0]);
            Assert.AreEqual(expCal, actCal[1]);
            Assert.AreEqual(expCal, actCal[2]);
            Assert.AreEqual(expCal, actCal[3]);
        }

        [Test]
        public void IsADateFormat()
        {
            // Cell content 2016-12-8 as an example
            // Cell show "12/8/2016"
            Assert.IsTrue(DateUtil.IsADateFormat(14, "m/d/yy"));
            // Cell show "Thursday, December 8, 2016"
            Assert.IsTrue(DateUtil.IsADateFormat(182, "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"));
            // Cell show "12/8"
            Assert.IsTrue(DateUtil.IsADateFormat(183, "m/d;@"));
            // Cell show "12/08/16"
            Assert.IsTrue(DateUtil.IsADateFormat(184, "mm/dd/yy;@"));
            // Cell show "8-Dec-16"
            Assert.IsTrue(DateUtil.IsADateFormat(185, "[$-409]d\\-mmm\\-yy;@"));
            // Cell show "D-16"
            Assert.IsTrue(DateUtil.IsADateFormat(186, "[$-409]mmmmm\\-yy;@"));

            // Cell show "2016年12月8日"
            Assert.IsTrue(DateUtil.IsADateFormat(165, "yyyy\"年\"m\"月\"d\"日\";@"));
            // Cell show "2016年12月"
            Assert.IsTrue(DateUtil.IsADateFormat(164, "yyyy\"年\"m\"月\";@"));
            // Cell show "12月8日"
            Assert.IsTrue(DateUtil.IsADateFormat(168, "m\"月\"d\"日\";@"));
            // Cell show "十二月八日"
            Assert.IsTrue(DateUtil.IsADateFormat(181, "[DBNum1][$-404]m\"月\"d\"日\";@"));
            // Cell show "贰零壹陆年壹拾贰月捌日"
            Assert.IsTrue(DateUtil.IsADateFormat(177, "[DBNum2][$-804]yyyy\"年\"m\"月\"d\"日\";@"));
            // Cell show "２０１６年１２月８日"
            Assert.IsTrue(DateUtil.IsADateFormat(178, "[DBNum3][$-804]yyyy\"年\"m\"月\"d\"日\";@"));
        }
    }
}

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
            TimeZone tz = LocaleUtil.GetUserTimeZone();
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
            TimeZone tz = LocaleUtil.GetUserTimeZone();
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
            TimeZone tz = LocaleUtil.GetUserTimeZone();
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
            TimeZone tz = LocaleUtil.GetUserTimeZone();
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
    }
}

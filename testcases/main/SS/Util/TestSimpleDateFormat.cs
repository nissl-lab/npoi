/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using System.Globalization;
using System.IO;
using NPOI.SS.Util;
using NPOI.Util;
using NUnit.Framework;

namespace TestCases.SS.Util
{
    [TestFixture]
    public class TestSimpleDateFormat
    {
        /*
         * Test SimpleDateFormat.Parse with ill-formed datetime string.
         */
        [Test]
        public void TestIllformedDatetimeParsing()
        {
            // See https://github.com/nissl-lab/npoi/issues/846 for more information.

            var format = new SimpleDateFormat();

            DateTime standardForm = default;
            DateTime illForm = default;

            Assert.DoesNotThrow(() => standardForm = format.Parse("2020-07-03T09:41:11-04:00"));
            Assert.DoesNotThrow(() => illForm = format.Parse("2020-07-03T 9:41:11-04:00"));
            Assert.Throws<FormatException>(() => format.Parse("2020-07-03T09: 1:11-04:00"));

            Assert.AreEqual(standardForm, illForm);
        }

    }

}
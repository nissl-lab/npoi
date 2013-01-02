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

using System;
using NPOI.HPSF;
using NPOI.HPSF.Wellknown;
using NPOI.Util;
using NUnit.Framework;

namespace TestCases.HPSF
{
    [TestFixture]
    public class TestVariantSupport
    {
        [Test]
        public void Test52337()
        {
            // document summary stream   from test1-excel.doc attached to Bugzilla 52337
            String hex =
                    "FE FF 00 00 05 01 02 00 00 00 00 00 00 00 00 00 00 00 00 00" +
                    "00 00 00 00 02 00 00 00 02 D5 CD D5 9C 2E 1B 10 93 97 08 00" +
                    "2B 2C F9 AE 44 00 00 00 05 D5 CD D5 9C 2E 1B 10 93 97 08 00" +
                    "2B 2C F9 AE 58 01 00 00 14 01 00 00 0C 00 00 00 01 00 00 00" +
                    "68 00 00 00 0F 00 00 00 70 00 00 00 05 00 00 00 98 00 00 00" +
                    "06 00 00 00 A0 00 00 00 11 00 00 00 A8 00 00 00 17 00 00 00" +
                    "B0 00 00 00 0B 00 00 00 B8 00 00 00 10 00 00 00 C0 00 00 00" +
                    "13 00 00 00 C8 00 00 00 16 00 00 00 D0 00 00 00 0D 00 00 00" +
                    "D8 00 00 00 0C 00 00 00 F3 00 00 00 02 00 00 00 E4 04 00 00" +
                    "1E 00 00 00 20 00 00 00 44 65 70 61 72 74 6D 65 6E 74 20 6F" +
                    "66 20 49 6E 74 65 72 6E 61 6C 20 41 66 66 61 69 72 73 00 00" +
                    "03 00 00 00 05 00 00 00 03 00 00 00 01 00 00 00 03 00 00 00" +
                    "47 03 00 00 03 00 00 00 0F 27 0B 00 0B 00 00 00 00 00 00 00" +
                    "0B 00 00 00 00 00 00 00 0B 00 00 00 00 00 00 00 0B 00 00 00" +
                    "00 00 00 00 1E 10 00 00 01 00 00 00 0F 00 00 00 54 68 69 73" +
                    "20 69 73 20 61 20 74 65 73 74 00 0C 10 00 00 02 00 00 00 1E" +
                    "00 00 00 06 00 00 00 54 69 74 6C 65 00 03 00 00 00 01 00 00" +
                    "00 00 00 00 A8 00 00 00 03 00 00 00 00 00 00 00 20 00 00 00" +
                    "01 00 00 00 38 00 00 00 02 00 00 00 40 00 00 00 01 00 00 00" +
                    "02 00 00 00 0C 00 00 00 5F 50 49 44 5F 48 4C 49 4E 4B 53 00" +
                    "02 00 00 00 E4 04 00 00 41 00 00 00 60 00 00 00 06 00 00 00" +
                    "03 00 00 00 74 00 74 00 03 00 00 00 09 00 00 00 03 00 00 00" +
                    "00 00 00 00 03 00 00 00 05 00 00 00 1F 00 00 00 0C 00 00 00" +
                    "4E 00 6F 00 72 00 6D 00 61 00 6C 00 32 00 2E 00 78 00 6C 00" +
                    "73 00 00 00 1F 00 00 00 0A 00 00 00 53 00 68 00 65 00 65 00" +
                    "74 00 31 00 21 00 44 00 32 00 00 00 ";
            byte[] bytes = HexRead.ReadFromString(hex);

            PropertySet ps = PropertySetFactory.Create(new ByteArrayInputStream(bytes));
            DocumentSummaryInformation dsi = (DocumentSummaryInformation)ps;
            Section s = (Section)dsi.Sections[0];

            Object hdrs = s.GetProperty(PropertyIDMap.PID_HEADINGPAIR);

            Assert.IsNotNull(hdrs, "PID_HEADINGPAIR was not found");

            Assert.IsTrue(hdrs is byte[], "PID_HEADINGPAIR: expected byte[] but was " + hdrs.GetType());
            // parse the value
            Vector v = new Vector((short)Variant.VT_VARIANT);
            v.Read((byte[])hdrs, 0);

            TypedPropertyValue[] items = v.Values;
            Assert.AreEqual(2, items.Length);

            Assert.IsNotNull(items[0].Value);
            Assert.IsTrue(items[0].Value is CodePageString, "first item must be CodePageString but was " + items[0].GetType());
            Assert.IsNotNull(items[1].Value);
            Assert.IsTrue(Number.IsInteger(items[1].Value), "second item must be Integer but was " + items[1].Value.GetType());
            Assert.AreEqual(1, (int)items[1].Value);

        }
    }
}
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

namespace TestCases.HSSF.Record.Pivot
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.PivotTable;
    using NPOI.Util;
    using TestCases.HSSF.Record;

    /**
     * Tests for {@link PageItemRecord}
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestPageItemRecord
    {

        [Test]
        public void TestMoreThanOneInfoItem_bug46917()
        {
            byte[] data = HexRead.ReadFromString("01 02 03 04 05 06 07 08 09 0A 0B 0C");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(PageItemRecord.sid, data);
            PageItemRecord rec = new PageItemRecord(in1);
            if (in1.Remaining == 6)
            {
                throw new AssertionException("Identified bug 46917");
            }
            Assert.AreEqual(0, in1.Remaining);

            Assert.AreEqual(4 + data.Length, rec.RecordSize);
        }

        [Test]
        public void TestSerialize()
        {
            ConfirmSerialize("01 02 03 04 05 06");
            ConfirmSerialize("01 02 03 04 05 06 07 08 09 0A 0B 0C");
            ConfirmSerialize("01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F 10 11 12");
        }

        private static PageItemRecord ConfirmSerialize(String hexDump)
        {
            byte[] data = HexRead.ReadFromString(hexDump);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(PageItemRecord.sid, data);
            PageItemRecord rec = new PageItemRecord(in1);
            Assert.AreEqual(0, in1.Remaining);
            Assert.AreEqual(4 + data.Length, rec.RecordSize);
            byte[] data2 = rec.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(PageItemRecord.sid, data, data2);
            return rec;
        }
    }

}
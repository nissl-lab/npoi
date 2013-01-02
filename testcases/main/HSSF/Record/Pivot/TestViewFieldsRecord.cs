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
    using NPOI.Util;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record.PivotTable;

    /**
     * Tests for {@link ViewFieldsRecord}
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestViewFieldsRecord
    {

        [Test]
        public void TestUnicodeFlag_bug46693()
        {
            byte[] data = HexRead.ReadFromString("01 00 01 00 01 00 04 00 05 00 00 6D 61 72 63 6F");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(ViewFieldsRecord.sid, data);
            ViewFieldsRecord rec = new ViewFieldsRecord(in1);
            if (in1.Remaining == 1)
            {
                throw new AssertionException("Identified bug 46693b");
            }
            Assert.AreEqual(0, in1.Remaining);
            Assert.AreEqual(4 + data.Length, rec.RecordSize);
        }

        [Test]
        public void TestSerialize()
        {
            // This hex data was produced by changing the 'Custom Name' property, 
            // available under 'Field Settings' from the 'PivotTable Field List' (Excel 2007)
            ConfirmSerialize("00 00 01 00 01 00 00 00 FF FF");
            ConfirmSerialize("01 00 01 00 01 00 04 00 05 00 00 6D 61 72 63 6F");
            ConfirmSerialize("01 00 01 00 01 00 04 00 0A 00 01 48 00 69 00 73 00 74 00 6F 00 72 00 79 00 2D 00 82 69 81 89");
        }

        private static ViewFieldsRecord ConfirmSerialize(String hexDump)
        {
            byte[] data = HexRead.ReadFromString(hexDump);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(ViewFieldsRecord.sid, data);
            ViewFieldsRecord rec = new ViewFieldsRecord(in1);
            Assert.AreEqual(0, in1.Remaining);
            Assert.AreEqual(4 + data.Length, rec.RecordSize);
            byte[] data2 = rec.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(ViewFieldsRecord.sid, data, data2);
            return rec;
        }
    }

}
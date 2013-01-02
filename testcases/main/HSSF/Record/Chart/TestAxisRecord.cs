
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */



namespace TestCases.HSSF.Record.Chart
{
    using System;
    using NPOI.HSSF.Record.Chart;
    using NUnit.Framework;
    using NPOI.HSSF.Record;

    /**
     * Tests the serialization and deserialization of the AxisRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestAxisRecord
    {
        byte[] data = new byte[] {
        (byte)0x00,(byte)0x00,                               // type
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00

        };

        public TestAxisRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            AxisRecord record = new AxisRecord(TestcaseRecordInputStream.Create((short)0x101d, data));
            Assert.AreEqual(AxisRecord.AXIS_TYPE_CATEGORY_OR_X_AXIS, record.AxisType);
            Assert.AreEqual(0, record.Reserved1);
            Assert.AreEqual(0, record.Reserved2);
            Assert.AreEqual(0, record.Reserved3);
            Assert.AreEqual(0, record.Reserved4);


            Assert.AreEqual(4 + 18, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            AxisRecord record = new AxisRecord();
            record.AxisType = (AxisRecord.AXIS_TYPE_CATEGORY_OR_X_AXIS);
            record.Reserved1 = (0);
            record.Reserved2 = (0);
            record.Reserved3 = (0);
            record.Reserved4 = (0);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}

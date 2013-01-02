
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
    using NPOI.HSSF.Record;
    using NUnit.Framework;
    using NPOI.HSSF.Record.Chart;
    /**
     * Tests the serialization and deserialization of the AxisParentRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestAxisParentRecord
    {
        byte[] data = new byte[] {
        (byte)0x00,(byte)0x00,                                   // axis type
        (byte)0x1D,(byte)0x02,(byte)0x00,(byte)0x00,             // x
        (byte)0xDD,(byte)0x00,(byte)0x00,(byte)0x00,             // y
        (byte)0x31,(byte)0x0B,(byte)0x00,(byte)0x00,             // width
        (byte)0x56,(byte)0x0B,(byte)0x00,(byte)0x00              // height
    };

        public TestAxisParentRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            AxisParentRecord record = new AxisParentRecord(TestcaseRecordInputStream.Create((short)0x1041, data));
            Assert.AreEqual(AxisParentRecord.AXIS_TYPE_MAIN, record.AxisType);
            Assert.AreEqual(0x021d, record.X);
            Assert.AreEqual(0xdd, record.Y);
            Assert.AreEqual(0x0b31, record.Width);
            Assert.AreEqual(0x0b56, record.Height);


            Assert.AreEqual(22, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            AxisParentRecord record = new AxisParentRecord();
            record.AxisType = (AxisParentRecord.AXIS_TYPE_MAIN);
            record.X = (0x021d);
            record.Y = (0xdd);
            record.Width = (0x0b31);
            record.Height = (0x0b56);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
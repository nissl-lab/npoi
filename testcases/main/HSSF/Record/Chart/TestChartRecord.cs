
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
    using NPOI.HSSF.Record.Chart;
    using NUnit.Framework;

    /**
     * Tests the serialization and deserialization of the ChartRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestChartRecord
    {
        byte[] data = new byte[] {
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0xE8,(byte)0xFF,(byte)0xD0,(byte)0x01,
        (byte)0xC8,(byte)0xCC,(byte)0xE5,(byte)0x00
    };

        public TestChartRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            ChartRecord record = new ChartRecord(TestcaseRecordInputStream.Create((short)0x1002,  data));
            Assert.AreEqual(0, record.X);
            Assert.AreEqual(0, record.Y);
            Assert.AreEqual(30474216, record.Width);
            Assert.AreEqual(15060168, record.Height);


            Assert.AreEqual(20, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            ChartRecord record = new ChartRecord();
            record.X=(0);
            record.Y=(0);
            record.Width=(30474216);
            record.Height=(15060168);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}

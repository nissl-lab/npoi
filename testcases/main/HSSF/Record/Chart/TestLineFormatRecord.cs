
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
     * Tests the serialization and deserialization of the LineFormatRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestLineFormatRecord
    {
        byte[] data = new byte[] {
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,    // colour
        (byte)0x00,(byte)0x00,                          // pattern
        (byte)0x00,(byte)0x00,                          // weight
        (byte)0x01,(byte)0x00,                          // format
        (byte)0x4D,(byte)0x00                           // index
    };

        public TestLineFormatRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            LineFormatRecord record = new LineFormatRecord(TestcaseRecordInputStream.Create((short)0x1007, data));
            Assert.AreEqual(0, record.LineColor);
            Assert.AreEqual(0, record.LinePattern);
            Assert.AreEqual(0, record.Weight);
            Assert.AreEqual(1, record.Format);
            Assert.AreEqual(true, record.IsAuto);
            Assert.AreEqual(false, record.IsDrawTicks);
            Assert.AreEqual(0x4d, record.ColourPaletteIndex);


            Assert.AreEqual(16, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            LineFormatRecord record = new LineFormatRecord();
            record.LineColor=(0);
            record.LinePattern=((short)0);
            record.Weight=((short)0);
            record.IsAuto=(true);
            record.IsDrawTicks=(false);
            record.ColourPaletteIndex=((short)0x4d);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}


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
     * Tests the serialization and deserialization of the AreaFormatRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestAreaFormatRecord
    {
        byte[] data = new byte[] {
        (byte)0xFF,(byte)0xFF,(byte)0xFF,(byte)0x00,    // forecolor
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,    // backcolor
        (byte)0x01,(byte)0x00,                          // pattern
        (byte)0x01,(byte)0x00,                          // format
        (byte)0x4E,(byte)0x00,                          // forecolor index
        (byte)0x4D,(byte)0x00                           // backcolor index

    };

        public TestAreaFormatRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            AreaFormatRecord record = new AreaFormatRecord(TestcaseRecordInputStream.Create((short)0x100a, data));
            Assert.AreEqual(0xFFFFFF, record.ForegroundColor);
            Assert.AreEqual(0x000000, record.BackgroundColor);
            Assert.AreEqual(1, record.Pattern);
            Assert.AreEqual(1, record.FormatFlags);
            Assert.AreEqual(true, record.IsAutomatic);
            Assert.AreEqual(false, record.IsInvert);
            Assert.AreEqual(0x4e, record.ForecolorIndex);
            Assert.AreEqual(0x4d, record.BackcolorIndex);


            Assert.AreEqual(20, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            AreaFormatRecord record = new AreaFormatRecord();
            record.ForegroundColor = (0xFFFFFF);
            record.BackgroundColor = (0x000000);
            record.Pattern = ((short)1);
            record.IsAutomatic = (true);
            record.IsInvert = (false);
            record.ForecolorIndex = ((short)0x4e);
            record.BackcolorIndex = ((short)0x4d);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
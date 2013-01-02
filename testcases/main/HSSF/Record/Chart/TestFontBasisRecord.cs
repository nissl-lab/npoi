
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
     * Tests the serialization and deserialization of the FontBasisRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestFontBasisRecord
    {
        byte[] data = new byte[] {
        (byte)0x28,(byte)0x1A,   // x basis
        (byte)0x9C,(byte)0x0F,   // y basis
        (byte)0xC8,(byte)0x00,   // height basis
        (byte)0x00,(byte)0x00,   // scale
        (byte)0x05,(byte)0x00    // index to font table
    };

        public TestFontBasisRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            FontBasisRecord record = new FontBasisRecord(TestcaseRecordInputStream.Create((short)0x1060, data));
            Assert.AreEqual(0x1a28, record.XBasis);
            Assert.AreEqual(0x0f9c, record.YBasis);
            Assert.AreEqual(0xc8, record.HeightBasis);
            Assert.AreEqual(0x00, record.Scale);
            Assert.AreEqual(0x05, record.IndexToFontTable);


            Assert.AreEqual(14, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            FontBasisRecord record = new FontBasisRecord();
            record.XBasis = ((short)0x1a28);
            record.YBasis = ((short)0x0f9c);
            record.HeightBasis = ((short)0xc8);
            record.Scale = ((short)0x00);
            record.IndexToFontTable = ((short)0x05);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
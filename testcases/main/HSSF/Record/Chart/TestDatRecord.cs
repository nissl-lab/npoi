
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
     * Tests the serialization and deserialization of the DatRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestDatRecord
    {
        byte[] data = new byte[] {
        (byte)0x0D,(byte)0x00   // options
    };

        public TestDatRecord()
        {

        }

        [Test]
        public void TestLoad()
        {

            DatRecord record = new DatRecord(TestcaseRecordInputStream.Create((short)0x1063,  data));
            Assert.AreEqual(0xD, record.Options);
            Assert.AreEqual(true, record.IsHorizontalBorder());
            Assert.AreEqual(false, record.IsVerticalBorder());
            Assert.AreEqual(true, record.IsBorder());
            Assert.AreEqual(true, record.IsShowSeriesKey());


            Assert.AreEqual(6, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            DatRecord record = new DatRecord();
            record.SetHorizontalBorder(true);
            record.SetVerticalBorder(false);
            record.SetBorder(true);
            record.SetShowSeriesKey(true);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
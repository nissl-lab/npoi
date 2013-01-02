
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
     * Tests the serialization and deserialization of the DataFormatRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestDataFormatRecord
    {
        byte[] data = new byte[] {
        (byte)0xFF,(byte)0xFF,      // point number
        (byte)0x00,(byte)0x00,      // series index
        (byte)0x00,(byte)0x00,      // series number
        (byte)0x00,(byte)0x00       // format flags
    };

        public TestDataFormatRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            DataFormatRecord record = new DataFormatRecord(TestcaseRecordInputStream.Create((short)0x1006, data));
            Assert.AreEqual(record.PointNumber, unchecked((short)0xFFFF));
            Assert.AreEqual(0, record.SeriesIndex);
            Assert.AreEqual(0, record.SeriesNumber);
            Assert.AreEqual(0, record.FormatFlags);
            Assert.AreEqual(false, record.UseExcel4Colors);


            Assert.AreEqual(12, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            DataFormatRecord record = new DataFormatRecord();
            record.PointNumber = (unchecked((short)0xFFFF));
            record.SeriesIndex = ((short)0);
            record.SeriesNumber = ((short)0);
            record.FormatFlags = ((short)0);
            record.UseExcel4Colors = (false);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
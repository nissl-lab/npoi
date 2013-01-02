
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
     * Tests the serialization and deserialization of the SeriesRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestSeriesRecord
    {
        byte[] data = new byte[] {
            (byte)0x01,(byte)0x00,      // category data type
            (byte)0x01,(byte)0x00,      // values data type
            (byte)0x1B,(byte)0x00,      // num categories
            (byte)0x1B,(byte)0x00,      // num values
            (byte)0x01,(byte)0x00,      // bubble series type
            (byte)0x00,(byte)0x00       // num bubble values
        };

        public TestSeriesRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            SeriesRecord record = new SeriesRecord(TestcaseRecordInputStream.Create(0x1003, data));
            Assert.AreEqual(SeriesRecord.CATEGORY_DATA_TYPE_NUMERIC, record.CategoryDataType);
            Assert.AreEqual(SeriesRecord.VALUES_DATA_TYPE_NUMERIC, record.ValuesDataType);
            Assert.AreEqual(27, record.NumCategories);
            Assert.AreEqual(27, record.NumValues);
            Assert.AreEqual(SeriesRecord.BUBBLE_SERIES_TYPE_NUMERIC, record.BubbleSeriesType);
            Assert.AreEqual(0, record.NumBubbleValues);


            Assert.AreEqual(16, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            SeriesRecord record = new SeriesRecord();
            record.CategoryDataType = (SeriesRecord.CATEGORY_DATA_TYPE_NUMERIC);
            record.ValuesDataType = (SeriesRecord.VALUES_DATA_TYPE_NUMERIC);
            record.NumCategories = ((short)27);
            record.NumValues = ((short)27);
            record.BubbleSeriesType = (SeriesRecord.BUBBLE_SERIES_TYPE_NUMERIC);
            record.NumBubbleValues = ((short)0);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
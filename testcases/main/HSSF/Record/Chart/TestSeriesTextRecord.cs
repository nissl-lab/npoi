
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
    using NPOI.Util;

    /**
     * Tests the serialization and deserialization of the SeriesTextRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    [TestFixture]
    public class TestSeriesTextRecord
    {

        private static byte[] SIMPLE_DATA = HexRead.ReadFromString("00 00 0C 00 56 61 6C 75 65 20 4E 75 6D 62 65 72");

        public TestSeriesTextRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            SeriesTextRecord record = new SeriesTextRecord(TestcaseRecordInputStream.Create(0x100d, SIMPLE_DATA));

            Assert.AreEqual((short)0, record.Id);
            Assert.AreEqual((byte)0x0C, record.Text.Length);

            Assert.AreEqual("Value Number", record.Text);


            Assert.AreEqual(SIMPLE_DATA.Length+4, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            SeriesTextRecord record = new SeriesTextRecord();



            record.Id = 0;
            record.Text = ("Value Number");


            byte[] recordBytes = record.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(SeriesTextRecord.sid, SIMPLE_DATA,recordBytes);
        }
    }
}
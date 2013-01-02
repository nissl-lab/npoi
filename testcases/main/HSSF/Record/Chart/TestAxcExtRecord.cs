
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
     * Tests the serialization and deserialization of the AxisOptionsRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Andrew C. Oliver(acoliver at apache.org)
     */
    [TestFixture]
    public class TestAxcExtRecord
    {
        byte[] data = new byte[] {        
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x01,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x01,(byte)0x00,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0x00,(byte)0xEF,(byte)0x00
    };

        public TestAxcExtRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            AxcExtRecord record = new AxcExtRecord(TestcaseRecordInputStream.Create((short)0x1062, data));
            Assert.AreEqual(0, record.MinimumDate);
            Assert.AreEqual(0, record.MaximumDate);
            Assert.AreEqual(1, record.MajorInterval);
            Assert.AreEqual(0, (short)record.MajorUnit);
            Assert.AreEqual(1, record.MinorInterval);
            Assert.AreEqual(0, (short)record.MinorUnit);
            Assert.AreEqual(0, (short)record.BaseUnit);
            Assert.AreEqual(0, record.CrossDate);
            Assert.AreEqual(239, record.Options);
            Assert.AreEqual(true, record.IsAutoMin);
            Assert.AreEqual(true, record.IsAutoMax);
            Assert.AreEqual(true, record.IsAutoMajor);
            Assert.AreEqual(true, record.IsAutoMinor);
            Assert.AreEqual(false, record.IsDateAxis);
            Assert.AreEqual(true, record.IsAutoBase);
            Assert.AreEqual(true, record.IsAutoCross);
            Assert.AreEqual(true, record.IsAutoDate);


            Assert.AreEqual(22, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            AxcExtRecord record = new AxcExtRecord();
            record.MinimumDate = ((short)0);
            record.MaximumDate = ((short)0);
            record.MajorInterval = ((short)1);
            record.MajorUnit = ((short)0);
            record.MinorInterval = ((short)1);
            record.MinorUnit = ((short)0);
            record.BaseUnit = ((short)0);
            record.CrossDate = ((short)0);
            record.Options = ((short)239);
            record.IsAutoMin = (true);
            record.IsAutoMax = (true);
            record.IsAutoMajor = (true);
            record.IsAutoMinor = (true);
            record.IsDateAxis = (false);
            record.IsAutoBase = (true);
            record.IsAutoCross = (true);
            record.IsAutoDate = (true);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
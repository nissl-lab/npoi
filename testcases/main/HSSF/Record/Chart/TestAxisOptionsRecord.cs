
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
    public class TestAxisOptionsRecord
    {
        byte[] data = new byte[] {        
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x01,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x01,(byte)0x00,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0x00,(byte)0xEF,(byte)0x00
    };

        public TestAxisOptionsRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            AxisOptionsRecord record = new AxisOptionsRecord(TestcaseRecordInputStream.Create((short)0x1062, data));
            Assert.AreEqual(0, record.MinimumCategory);
            Assert.AreEqual(0, record.MaximumCategory);
            Assert.AreEqual(1, record.MajorUnitValue);
            Assert.AreEqual(0, record.MajorUnit);
            Assert.AreEqual(1, record.MinorUnitValue);
            Assert.AreEqual(0, record.MinorUnit);
            Assert.AreEqual(0, record.BaseUnit);
            Assert.AreEqual(0, record.CrossingPoint);
            Assert.AreEqual(239, record.Options);
            Assert.AreEqual(true, record.IsDefaultMinimum);
            Assert.AreEqual(true, record.IsDefaultMaximum);
            Assert.AreEqual(true, record.IsDefaultMajor);
            Assert.AreEqual(true, record.IsDefaultMinorUnit);
            Assert.AreEqual(false, record.IsDate);
            Assert.AreEqual(true, record.IsDefaultBase);
            Assert.AreEqual(true, record.IsDefaultCross);
            Assert.AreEqual(true, record.IsDefaultDateSettings);


            Assert.AreEqual(22, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            AxisOptionsRecord record = new AxisOptionsRecord();
            record.MinimumCategory = ((short)0);
            record.MaximumCategory = ((short)0);
            record.MajorUnitValue = ((short)1);
            record.MajorUnit = ((short)0);
            record.MinorUnitValue = ((short)1);
            record.MinorUnit = ((short)0);
            record.BaseUnit = ((short)0);
            record.CrossingPoint = ((short)0);
            record.Options = ((short)239);
            record.IsDefaultMinimum = (true);
            record.IsDefaultMaximum = (true);
            record.IsDefaultMajor = (true);
            record.IsDefaultMinorUnit = (true);
            record.IsDate = (false);
            record.IsDefaultBase = (true);
            record.IsDefaultCross = (true);
            record.IsDefaultDateSettings = (true);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
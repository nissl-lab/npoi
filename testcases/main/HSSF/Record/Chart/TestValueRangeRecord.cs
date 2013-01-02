
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
     * Tests the serialization and deserialization of the ValueRangeRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestValueRangeRecord
    {
        byte[] data = new byte[] {
            (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,     // min axis value
            (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,     // max axis value
            (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,     // major increment
            (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,     // minor increment
            (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,     // cross over
            (byte)0x1F,(byte)0x01                                    // options
         };

        public TestValueRangeRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            ValueRangeRecord record = new ValueRangeRecord(TestcaseRecordInputStream.Create(0x101f, data));
            Assert.AreEqual(0.0, record.MinimumAxisValue, 0.001);
            Assert.AreEqual(0.0, record.MaximumAxisValue, 0.001);
            Assert.AreEqual(0.0, record.MajorIncrement, 0.001);
            Assert.AreEqual(0.0, record.MinorIncrement, 0.001);
            Assert.AreEqual(0.0, record.CategoryAxisCross, 0.001);
            Assert.AreEqual(0x011f, record.Options);
            Assert.AreEqual(true, record.IsAutomaticMinimum);
            Assert.AreEqual(true, record.IsAutomaticMaximum);
            Assert.AreEqual(true, record.IsAutomaticMajor);
            Assert.AreEqual(true, record.IsAutomaticMinor);
            Assert.AreEqual(true, record.IsAutomaticCategoryCrossing);
            Assert.AreEqual(false, record.IsLogarithmicScale);
            Assert.AreEqual(false, record.IsValuesInReverse);
            Assert.AreEqual(false, record.IsCrossCategoryAxisAtMaximum);
            Assert.AreEqual(true, record.IsReserved);

            Assert.AreEqual(42 + 4, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            ValueRangeRecord record = new ValueRangeRecord();
            record.MinimumAxisValue=(0);
            record.MaximumAxisValue=(0);
            record.MajorIncrement=(0);
            record.MinorIncrement=(0);
            record.CategoryAxisCross=(0);
            record.IsAutomaticMinimum=(true);
            record.IsAutomaticMaximum=(true);
            record.IsAutomaticMajor=(true);
            record.IsAutomaticMinor=(true);
            record.IsAutomaticCategoryCrossing=(true);
            record.IsLogarithmicScale=(false);
            record.IsValuesInReverse=(false);
            record.IsCrossCategoryAxisAtMaximum=(false);
            record.IsReserved=(true);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
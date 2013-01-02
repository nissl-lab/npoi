
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
     * Tests the serialization and deserialization of the SeriesLabelsRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestAttachedLabelRecord
    {
        byte[] data = new byte[] {
        (byte)0x03,(byte)0x00
    };

        public TestAttachedLabelRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            AttachedLabelRecord record = new AttachedLabelRecord(TestcaseRecordInputStream.Create(0x100c, data));
            Assert.AreEqual(3, record.FormatFlags);
            Assert.AreEqual(true, record.IsShowActual);
            Assert.AreEqual(true, record.IsShowPercent);
            Assert.AreEqual(false, record.IsLabelAsPercentage);
            Assert.AreEqual(false, record.IsSmoothedLine);
            Assert.AreEqual(false, record.IsShowLabel);
            Assert.AreEqual(false, record.IsShowBubbleSizes);


            Assert.AreEqual(2 + 4, record.RecordSize);

        }
        [Test]
        public void TestStore()
        {
            AttachedLabelRecord record = new AttachedLabelRecord();
            record.IsShowActual=(true);
            record.IsShowPercent=(true);
            record.IsLabelAsPercentage = (false);
            record.IsSmoothedLine = (false);
            record.IsShowLabel = (false);
            record.IsShowBubbleSizes = (false);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
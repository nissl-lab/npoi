
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
     * Tests the serialization and deserialization of the BarRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestBarRecord
    {
        byte[] data = new byte[] {
        (byte)0x00,(byte)0x00,   // bar space
        (byte)0x96,(byte)0x00,   // category space
        (byte)0x00,(byte)0x00    // format flags
    };

        public TestBarRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            BarRecord record = new BarRecord(TestcaseRecordInputStream.Create((short)0x1017, data));
            Assert.AreEqual(0, record.BarSpace);
            Assert.AreEqual(0x96, record.CategorySpace);
            Assert.AreEqual(0, record.FormatFlags);
            Assert.AreEqual(false, record.IsHorizontal);
            Assert.AreEqual(false, record.IsStacked);
            Assert.AreEqual(false, record.IsDisplayAsPercentage);
            Assert.AreEqual(false, record.IsShadow);


            Assert.AreEqual(10, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            BarRecord record = new BarRecord();
            record.BarSpace = ((short)0);
            record.CategorySpace = ((short)0x96);
            record.IsHorizontal = (false);
            record.IsStacked = (false);
            record.IsDisplayAsPercentage = (false);
            record.IsShadow = (false);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
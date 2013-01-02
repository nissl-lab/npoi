
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
     * Tests the serialization and deserialization of the AreaRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestAreaRecord
    {
        byte[] data = new byte[] {
        (byte)0x02,(byte)0x00       // format flags
    };

        public TestAreaRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            AreaRecord record = new AreaRecord(TestcaseRecordInputStream.Create((short)0x101A, data));
            Assert.AreEqual(2, record.FormatFlags);
            Assert.AreEqual(false, record.IsStacked);
            Assert.AreEqual(true, record.IsDisplayAsPercentage);
            Assert.AreEqual(false, record.IsShadow);


            Assert.AreEqual(6, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            AreaRecord record = new AreaRecord();
            record.IsStacked = (false);
            record.IsDisplayAsPercentage = (true);
            record.IsShadow = (false);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
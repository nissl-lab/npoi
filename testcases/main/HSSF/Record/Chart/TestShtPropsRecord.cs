
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
     * Tests the serialization and deserialization of the SheetPropertiesRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestShtPropsRecord
    {
        byte[] data = new byte[] {
        (byte)0x0A,(byte)0x00,
        (byte)0x00,
        (byte)0x00       // not sure where that last byte comes from
        };

        public TestShtPropsRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            ShtPropsRecord record = new ShtPropsRecord(TestcaseRecordInputStream.Create(0x1044, data));
            Assert.AreEqual(10, record.Flags);
            Assert.AreEqual(false, record.IsManSerAlloc);
            Assert.AreEqual(true, record.IsPlotVisibleOnly);
            Assert.AreEqual(false, record.IsNotSizeWithWindow);
            Assert.AreEqual(true, record.IsManPlotArea);
            Assert.AreEqual(false, record.IsAlwaysAutoPlotArea);
            Assert.AreEqual(0, record.Blank);


            Assert.AreEqual(8, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            ShtPropsRecord record = new ShtPropsRecord();
            record.IsManSerAlloc=(false);
            record.IsPlotVisibleOnly=(true);
            record.IsNotSizeWithWindow=(false);
            record.IsManPlotArea=(true);
            record.IsAlwaysAutoPlotArea=(false);
            record.Blank=0;


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
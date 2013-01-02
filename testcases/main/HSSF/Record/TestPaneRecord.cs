
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



namespace TestCases.HSSF.Record
{
    using System;
    using NPOI.HSSF.Record;

    using NUnit.Framework;

    /**
     * Tests the serialization and deserialization of the PaneRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestPaneRecord
    {
        byte[] data = new byte[] {
        (byte)0x01, (byte)0x00,
        (byte)0x02, (byte)0x00,
        (byte)0x03, (byte)0x00,
        (byte)0x04, (byte)0x00,
        (byte)0x02, (byte)0x00,
    };

        public TestPaneRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            PaneRecord record = new PaneRecord(TestcaseRecordInputStream.Create((short)0x41, data));


            Assert.AreEqual((short)1, record.X);
            Assert.AreEqual((short)2, record.Y);
            Assert.AreEqual((short)3, record.TopRow);
            Assert.AreEqual((short)4, record.LeftColumn);
            Assert.AreEqual(PaneRecord.ACTIVE_PANE_LOWER_LEFT, record.ActivePane);

            Assert.AreEqual(14, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            PaneRecord record = new PaneRecord();

            record.X=((short)1);
            record.Y=((short)2);
            record.TopRow=((short)3);
            record.LeftColumn=((short)4);
            record.ActivePane=(PaneRecord.ACTIVE_PANE_LOWER_LEFT);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}
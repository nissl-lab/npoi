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
     * Tests the serialization and deserialization of the NoteRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestNoteStructureSubRecord
    {
        private byte[] data = new byte[] {
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, (byte)0x80, 0x00, 0x00, 0x00,
            0x00, 0x00, (byte)0xBF, 0x00, 0x00, 0x00, 0x00, 0x00, (byte)0x81, 0x01,
            (byte)0xCC, (byte)0xEC
        };

        public TestNoteStructureSubRecord()
        {

        }
        [Test]
        public void TestRead()
        {

            NoteStructureSubRecord record = new NoteStructureSubRecord(TestcaseRecordInputStream.Create(NoteStructureSubRecord.sid, data),data.Length);

            Assert.AreEqual(NoteStructureSubRecord.sid, record.Sid);
            Assert.AreEqual(data.Length , record.DataSize);

        }
        [Test]
        public void TestWrite()
        {
            NoteStructureSubRecord record = new NoteStructureSubRecord();
            Assert.AreEqual(NoteStructureSubRecord.sid, record.Sid);
            Assert.AreEqual(data.Length , record.DataSize);

            byte[] ser = record.Serialize();
            Assert.AreEqual(ser.Length - 4, data.Length);

        }
        [Test]
        public void TestClone()
        {
            NoteStructureSubRecord record = new NoteStructureSubRecord();
            byte[] src = record.Serialize();

            NoteStructureSubRecord cloned = (NoteStructureSubRecord)record.Clone();
            byte[] cln = cloned.Serialize();

            Assert.AreEqual(record.DataSize, cloned.DataSize);
            Assert.IsTrue(NPOI.Util.Arrays.Equals(src, cln));
        }
    }
}
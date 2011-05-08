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
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Tests the serialization and deserialization of the NoteRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     * @author Yegor Kozlov
     */
    [TestClass]
    public class TestNoteRecord
    {
        private byte[] data = new byte[] {
        0x06, 0x00, 0x01, 0x00, 0x02, 0x00, 0x02, 0x04, 0x1A, 0x00,
        0x00, 0x41, 0x70, 0x61, 0x63, 0x68, 0x65, 0x20, 0x53, 0x6F,
        0x66, 0x74, 0x77, 0x61, 0x72, 0x65, 0x20, 0x46, 0x6F, 0x75,
        0x6E, 0x64, 0x61, 0x74, 0x69, 0x6F, 0x6E, 0x00
    };

        public TestNoteRecord()
        {

        }
        [TestMethod]
        public void TestRead()
        {

            NoteRecord record = new NoteRecord(TestcaseRecordInputStream.Create(NoteRecord.sid, data));

            Assert.AreEqual(NoteRecord.sid, record.Sid);

            Assert.AreEqual(6, record.Row);
            Assert.AreEqual(1, record.Column);
            Assert.AreEqual(NoteRecord.NOTE_VISIBLE, record.Flags);
            Assert.AreEqual(1026, record.ShapeId);
            Assert.AreEqual("Apache Software Foundation", record.Author);

        }
        [TestMethod]
        public void TestWrite()
        {
            NoteRecord record = new NoteRecord();
            Assert.AreEqual(NoteRecord.sid, record.Sid);

            record.Row = ((short)6);
            record.Column = ((short)1);
            record.Flags = (NoteRecord.NOTE_VISIBLE);
            record.ShapeId = ((short)1026);
            record.Author = ("Apache Software Foundation");

            byte[] ser = record.Serialize();
            Assert.AreEqual(ser.Length - 4, data.Length);

            byte[] recdata = new byte[ser.Length - 4];
            System.Array.Copy(ser, 4, recdata, 0, recdata.Length);
            Assert.IsTrue(NPOI.Util.Arrays.Equals(data, recdata));
        }
        [TestMethod]
        public void TestClone()
        {
            NoteRecord record = new NoteRecord();

            record.Row = ((short)1);
            record.Column = ((short)2);
            record.Flags = (NoteRecord.NOTE_VISIBLE);
            record.ShapeId = ((short)1026);
            record.Author = ("Apache Software Foundation");

            NoteRecord cloned = (NoteRecord)record.Clone();
            Assert.AreEqual(record.Row, cloned.Row);
            Assert.AreEqual(record.Column, cloned.Column);
            Assert.AreEqual(record.Flags, cloned.Flags);
            Assert.AreEqual(record.ShapeId, cloned.ShapeId);
            Assert.AreEqual(record.Author, cloned.Author);

            //finally check that the Serialized data is1 the same
            byte[] src = record.Serialize();
            byte[] cln = cloned.Serialize();
            Assert.IsTrue(NPOI.Util.Arrays.Equals(src, cln));
        }
    }
}
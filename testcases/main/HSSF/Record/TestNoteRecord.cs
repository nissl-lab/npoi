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
    using NPOI.Util;

    /**
     * Tests the serialization and deserialization of the NoteRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestNoteRecord
    {
        private byte[] testData = HexRead.ReadFromString(
            "06 00 01 00 02 00 02 04 " +
            "1A 00 00 " +
            "41 70 61 63 68 65 20 53 6F 66 74 77 61 72 65 20 46 6F 75 6E 64 61 74 69 6F 6E " +
            "00" // padding byte
            );

        [Test]
        public void TestRead()
        {

            NoteRecord record = new NoteRecord(TestcaseRecordInputStream.Create(NoteRecord.sid, testData));

            Assert.AreEqual(NoteRecord.sid, record.Sid);

            Assert.AreEqual(6, record.Row);
            Assert.AreEqual(1, record.Column);
            Assert.AreEqual(NoteRecord.NOTE_VISIBLE, record.Flags);
            Assert.AreEqual(1026, record.ShapeId);
            Assert.AreEqual("Apache Software Foundation", record.Author);

        }
        [Test]
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
            TestcaseRecordInputStream.ConfirmRecordEncoding(NoteRecord.sid, testData, ser);
        }
        [Test]
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
        [Test]
        public void TestUnicodeAuthor()
        {
            // This sample data was created by setting the 'user name' field in the 'Personalize' 
            // section of Excel's options to \u30A2\u30D1\u30C3\u30C1\u65CF, and then 
            // creating a cell comment.
            byte[] data = HexRead.ReadFromString("01 00 01 00 00 00 03 00 " +
                    "05 00 01 " + // len=5, 16bit
                    "A2 30 D1 30 C3 30 C1 30 CF 65 " + // character data 
                    "00 " // padding byte
                    );
            RecordInputStream in1 = TestcaseRecordInputStream.Create(NoteRecord.sid, data);
            NoteRecord nr = new NoteRecord(in1);
            if ("\u00A2\u0030\u00D1\u0030\u00C3".Equals(nr.Author))
            {
                throw new AssertionException("Identified bug in reading note with unicode author");
            }
            Assert.AreEqual("\u30A2\u30D1\u30C3\u30C1\u65CF", nr.Author);
            Assert.IsTrue(nr.AuthorIsMultibyte);

            byte[] ser = nr.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(NoteRecord.sid, data, ser);

            // Re-check
            in1 = TestcaseRecordInputStream.Create(ser);
            nr = new NoteRecord(in1);
            Assert.AreEqual("\u30A2\u30D1\u30C3\u30C1\u65CF", nr.Author);
            Assert.IsTrue(nr.AuthorIsMultibyte);


            // Change to a non unicode author, will stop being unicode
            nr.Author = ("Simple");
            ser = nr.Serialize();
            in1 = TestcaseRecordInputStream.Create(ser);
            nr = new NoteRecord(in1);

            Assert.AreEqual("Simple", nr.Author);
            Assert.IsFalse(nr.AuthorIsMultibyte);

            // Now set it back again
            nr.Author = ("Unicode\u1234");
            ser = nr.Serialize();
            in1 = TestcaseRecordInputStream.Create(ser);
            nr = new NoteRecord(in1);

            Assert.AreEqual("Unicode\u1234", nr.Author);
            Assert.IsTrue(nr.AuthorIsMultibyte);
        }
    }
}
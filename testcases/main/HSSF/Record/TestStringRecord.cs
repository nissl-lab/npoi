/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record
{
    using System;
    using System.Text;
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * Tests the serialization and deserialization of the StringRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     */
    [TestFixture]
    public class TestStringRecord
    {
        private static byte[] data = HexRead.ReadFromString(
                "0B 00 " + // length
                "00 " +    // option
                // string
                "46 61 68 72 7A 65 75 67 74 79 70"
        );

        [Test]
        public void TestLoad()
        {
            StringRecord record = new StringRecord(TestcaseRecordInputStream.Create(0x207, data));
            Assert.AreEqual("Fahrzeugtyp", record.String);

            Assert.AreEqual(18, record.RecordSize);
        }

        [Test]
        public void TestStore()
        {
            StringRecord record = new StringRecord();
            record.String = (/*setter*/"Fahrzeugtyp");

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
            {
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
            }
        }

        [Test]
        public void TestContinue()
        {
            int MAX_BIFF_DATA = RecordInputStream.MAX_RECORD_DATA_SIZE;
            int TEXT_LEN = MAX_BIFF_DATA + 1000; // deliberately over-size
            String textChunk = "ABCDEGGHIJKLMNOP"; // 16 chars
            StringBuilder sb = new StringBuilder(16384);
            while (sb.Length < TEXT_LEN)
            {
                sb.Append(textChunk);
            }
            sb.Length = (/*setter*/TEXT_LEN);

            StringRecord sr = new StringRecord();
            sr.String = (/*setter*/sb.ToString());
            byte[] ser = sr.Serialize();
            Assert.AreEqual(StringRecord.sid, LittleEndian.GetUShort(ser, 0));
            if (LittleEndian.GetUShort(ser, 2) > MAX_BIFF_DATA)
            {
                Assert.Fail("StringRecord should have been split with a continue record");
            }

            // Confirm expected size of first record, and ushort strLen.
            Assert.AreEqual(MAX_BIFF_DATA, LittleEndian.GetUShort(ser, 2));
            Assert.AreEqual(TEXT_LEN, LittleEndian.GetUShort(ser, 4));

            // Confirm first few bytes of ContinueRecord
            LittleEndianByteArrayInputStream crIn = new LittleEndianByteArrayInputStream(ser, (MAX_BIFF_DATA + 4));
            int nCharsInFirstRec = MAX_BIFF_DATA - (2 + 1); // strLen, optionFlags
            int nCharsInSecondRec = TEXT_LEN - nCharsInFirstRec;
            Assert.AreEqual(ContinueRecord.sid, crIn.ReadUShort());
            Assert.AreEqual(1 + nCharsInSecondRec, crIn.ReadUShort());
            Assert.AreEqual(0, crIn.ReadUByte());
            Assert.AreEqual('N', crIn.ReadUByte());
            Assert.AreEqual('O', crIn.ReadUByte());

            // re-read and make sure string value is the same
            RecordInputStream in1 = TestcaseRecordInputStream.Create(ser);
            StringRecord sr2 = new StringRecord(in1);
            Assert.AreEqual(sb.ToString(), sr2.String);
        }
    }

}

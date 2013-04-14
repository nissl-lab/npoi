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
    using System.Collections;
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * Tests the serialization and deserialization of the ObjRecord class works correctly.
     * Test data taken directly from a real Excel file.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestObjRecord
    {
        /**
         * OBJ record data containing two sub-records.
         * The data taken directly from a real Excel file.
         *
         * [OBJ]
         *     [ftCmo]
         *     [ftEnd]
         */
        private static byte[] recdata = {
            0x15, 0x00, 0x12, 0x00, 0x06, 0x00, 0x01, 0x00, 0x11, 0x60,
            (byte)0xF4, 0x02, 0x41, 0x01, 0x14, 0x10, 0x1F, 0x02, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00
            // TODO - this data seems to require two extra bytes padding. not sure where original file is.
            // it's not bug 38607 attachment 17639
        };

        private static byte[] recdataNeedingPadding = {
    	    21, 0, 18, 0, 0, 0, 1, 0, 17, 96, 0, 0, 0, 0, 56, 111, unchecked((byte)-52), 3, 0, 0, 0, 0, 6, 0, 2, 0, 0, 0, 0, 0, 0, 0
        };
        [Test]
        public void TestLoad()
        {
            ObjRecord record = new ObjRecord(TestcaseRecordInputStream.Create(ObjRecord.sid, recdata));

            Assert.AreEqual(26, record.RecordSize - 4);

            IList subrecords = record.SubRecords;
            Assert.AreEqual(2, subrecords.Count);
            Assert.IsTrue(subrecords[0] is CommonObjectDataSubRecord);
            Assert.IsTrue(subrecords[1] is EndSubRecord);

        }
        [Test]
        public void TestStore()
        {
            ObjRecord record = new ObjRecord(TestcaseRecordInputStream.Create(ObjRecord.sid, recdata));

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(26, recordBytes.Length - 4);
            byte[] subData = new byte[recdata.Length];
            System.Array.Copy(recordBytes, 4, subData, 0, subData.Length);
            Assert.IsTrue(NPOI.Util.Arrays.Equals(recdata, subData));
        }
        [Test]
        public void TestConstruct()
        {
            ObjRecord record = new ObjRecord();
            CommonObjectDataSubRecord ftCmo = new CommonObjectDataSubRecord();
            ftCmo.ObjectType = (CommonObjectType.Comment);
            ftCmo.ObjectId = ((short)1024);
            ftCmo.IsLocked = (true);
            ftCmo.IsPrintable = (true);
            ftCmo.IsAutoFill = (true);
            ftCmo.IsAutoline = (true);
            record.AddSubRecord(ftCmo);
            EndSubRecord ftEnd = new EndSubRecord();
            record.AddSubRecord(ftEnd);

            //Serialize and Read again
            byte[] recordBytes = record.Serialize();
            //cut off the record header
            byte[] bytes = new byte[recordBytes.Length - 4];
            System.Array.Copy(recordBytes, 4, bytes, 0, bytes.Length);

            record = new ObjRecord(TestcaseRecordInputStream.Create(ObjRecord.sid, bytes));
            IList subrecords = record.SubRecords;
            Assert.AreEqual(2, subrecords.Count);
            Assert.IsTrue(subrecords[0] is CommonObjectDataSubRecord);
            Assert.IsTrue(subrecords[1] is EndSubRecord);
        }
        [Test]
        public void TestReadWriteWithPadding_bug45133()
        {
            ObjRecord record = new ObjRecord(TestcaseRecordInputStream.Create(ObjRecord.sid, recdataNeedingPadding));

            if (record.RecordSize == 34)
            {
                throw new AssertionException("Identified bug 45133");
            }

            Assert.AreEqual(36, record.RecordSize);

            IList subrecords = record.SubRecords;
            Assert.AreEqual(3, subrecords.Count);
            Assert.AreEqual(typeof(CommonObjectDataSubRecord), subrecords[0].GetType());
            Assert.AreEqual(typeof(GroupMarkerSubRecord), subrecords[1].GetType());
            Assert.AreEqual(typeof(EndSubRecord), subrecords[2].GetType());
        }
        /**
 * Check that ObjRecord tolerates and preserves padding to a 4-byte boundary
 * (normally padding is to a 2-byte boundary).
 */
        [Test]
        public void Test4BytePadding()
        {
            // actual data from file saved by Excel 2007
            byte[] data = HexRead.ReadFromString(""
                    + "15 00 12 00  1E 00 01 00  11 60 B4 6D  3C 01 C4 06 "
                    + "49 06 00 00  00 00 00 00  00 00 00 00");
            // this data seems to have 2 extra bytes of padding more than usual
            // the total may have been padded to the nearest quad-byte length
            RecordInputStream in1 = TestcaseRecordInputStream.Create(ObjRecord.sid, data);
            // check read OK
            ObjRecord record = new ObjRecord(in1);
            // check that it re-serializes to the same data
            byte[] ser = record.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(ObjRecord.sid, data, ser);
        }
    }
}
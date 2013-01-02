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
    using System.IO;
    using NUnit.Framework;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;
    using NPOI.Util;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;

    /**
     * Tests the serialization and deserialization of the LbsDataSubRecord class works correctly.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestLbsDataSubRecord
    {

        /**
         * Test Read-write round trip
         * Test data was taken from 47701.xls
         */
        [Test]
        public void Test_47701()
        {
            byte[] data = HexRead.ReadFromString(
                            "15, 00, 12, 00, 12, 00, 02, 00, 11, 20, " +
                            "00, 00, 00, 00, 80, 3D, 03, 05, 00, 00, " +
                            "00, 00, 0C, 00, 14, 00, 00, 00, 00, 00, " +
                            "00, 00, 00, 00, 00, 00, 01, 00, 0A, 00, " +
                            "00, 00, 10, 00, 01, 00, 13, 00, EE, 1F, " +
                            "10, 00, 09, 00, 40, 9F, 74, 01, 25, 09, " +
                            "00, 0C, 00, 07, 00, 07, 00, 07, 04, 00, " +
                            "00, 00, 08, 00, 00, 00");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(ObjRecord.sid, data);
            // check read OK
            ObjRecord record = new ObjRecord(in1);
            Assert.AreEqual(3, record.SubRecords.Count);
            SubRecord sr = record.SubRecords[(2)];
            Assert.IsTrue(sr is LbsDataSubRecord);
            LbsDataSubRecord lbs = (LbsDataSubRecord)sr;
            Assert.AreEqual(4, lbs.NumberOfItems);

            Assert.IsTrue(lbs.Formula is AreaPtg);
            AreaPtg ptg = (AreaPtg)lbs.Formula;
            CellRangeAddress range = new CellRangeAddress(
                    ptg.FirstRow, ptg.LastRow, ptg.FirstColumn, ptg.LastColumn);
            Assert.AreEqual("H10:H13", range.FormatAsString());

            // check that it re-Serializes to the same data
            byte[] ser = record.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(ObjRecord.sid, data, ser);
        }

        /**
         * Test data was taken from the file attached to Bugzilla 45778
         */
        [Test]
        public void Test_45778()
        {
            byte[] data = HexRead.ReadFromString(
                                    "15, 00, 12, 00, 14, 00, 01, 00, 01, 00, " +
                                    "01, 21, 00, 00, 3C, 13, F4, 03, 00, 00, " +
                                    "00, 00, 0C, 00, 14, 00, 00, 00, 00, 00, " +
                                    "00, 00, 00, 00, 00, 00, 01, 00, 08, 00, 00, " +
                                    "00, 10, 00, 00, 00, " +
                                     "13, 00, EE, 1F, " +
                                     "00, 00, " +
                                     "08, 00, " +  //number of items
                                     "08, 00, " + //selected item
                                     "01, 03, " + //flags
                                     "00, 00, " + //objId
                //LbsDropData
                                     "0A, 00, " + //flags
                                     "14, 00, " + //the number of lines to be displayed in the dropdown
                                     "6C, 00, " + //the smallest width in pixels allowed for the dropdown window
                                     "00, 00, " +  //num chars
                                     "00, 00");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(ObjRecord.sid, data);
            // check read OK
            ObjRecord record = new ObjRecord(in1);

            SubRecord sr = record.SubRecords[(2)];
            Assert.IsTrue(sr is LbsDataSubRecord);
            LbsDataSubRecord lbs = (LbsDataSubRecord)sr;
            Assert.AreEqual(8, lbs.NumberOfItems);
            Assert.IsNull(lbs.Formula);

            // check that it re-Serializes to the same data
            byte[] ser = record.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(ObjRecord.sid, data, ser);

        }

        /**
         * Test data produced by OpenOffice 3.1 by opening  and saving 47701.xls
         * There are 5 pAdding bytes that are Removed by POI
         */
        [Test]
        public void Test_Remove_padding()
        {
            byte[] data = HexRead.ReadFromString(
                            "5D, 00, 4C, 00, " +
                            "15, 00, 12, 00, 12, 00, 01, 00, 11, 00, " +
                            "00, 00, 00, 00, 00, 00, 00, 00, 00, 00, " +
                            "00, 00, 0C, 00, 14, 00, 00, 00, 00, 00, " +
                            "00, 00, 00, 00, 00, 00, 01, 00, 09, 00, " +
                            "00, 00, 0F, 00, 01, 00, " +
                            "13, 00, 1B, 00, " +
                            "10, 00, " + //next 16 bytes is a ptg aray
                            "09, 00, 00, 00, 00, 00, 25, 09, 00, 0C, 00, 07, 00, 07, 00, 00, " +
                            "01, 00, " + //num lines
                            "00, 00, " + //selected
                             "08, 00, " +
                             "00, 00, " +
                             "00, 00, 00, 00, 00"); //pAdding bytes

            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
            // check read OK
            ObjRecord record = new ObjRecord(in1);
            // check that it re-Serializes to the same data
            byte[] ser = record.Serialize();

            Assert.AreEqual(data.Length - 5, ser.Length);
            for (int i = 0; i < ser.Length; i++) Assert.AreEqual(data[i], ser[i]);

            //check we can read the Trimmed record
            RecordInputStream in2 = TestcaseRecordInputStream.Create(ser);
            ObjRecord record2 = new ObjRecord(in2);
            byte[] ser2 = record2.Serialize();
            Assert.IsTrue(Arrays.Equals(ser, ser2));
        }
        [Test]
        public void Test_LbsDropData()
        {
            byte[] data = HexRead.ReadFromString(
                //LbsDropData
                                     "0A, 00, " + //flags
                                     "14, 00, " + //the number of lines to be displayed in the dropdown
                                     "6C, 00, " + //the smallest width in pixels allowed for the dropdown window
                                     "00, 00, " +  //num chars
                                     "00, " +      //compression flag
                                     "00");        //pAdding byte

            LittleEndianInputStream in1 = new LittleEndianInputStream(new MemoryStream(data));

            LbsDropData lbs = new LbsDropData(in1);

            MemoryStream baos = new MemoryStream();
            lbs.Serialize(new LittleEndianOutputStream(baos));

            Assert.IsTrue(Arrays.Equals(data, baos.ToArray()));
        }
    }

}
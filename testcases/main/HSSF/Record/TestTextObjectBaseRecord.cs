
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
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    /**
     * Tests the serialization and deserialization of the TextObjectBaseRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestTextObjectBaseRecord
    {
	    /** data for one TXO rec and two continue recs */
        private static byte[] data = HexRead.ReadFromString(
            "B6 01 " + // TextObjectRecord.sid
            "12 00 " + // size 18
            "44 02 02 00 00 00 00 00" +
            "00 00 " +
            "02 00 " + // strLen 2
            "10 00 " + // 16 bytes for 2 format runs
            "00 00" +
            "00 00 " +
            "3C 00 " + // ContinueRecord.sid
            "03 00 " + // size 3
            "00 " + // unicode compressed
            "41 42 " + // 'AB'
            "3C 00 " + // ContinueRecord.sid
            "10 00 " + // size 16
            "00 00 18 00 00 00 00 00 " +
            "02 00 00 00 00 00 00 00 "
        );

        [Test]
        public void TestLoad()
        {
            TextObjectRecord record = new TextObjectRecord(TestcaseRecordInputStream.Create(data));
            Assert.AreEqual(HorizontalTextAlignment.Center, record.HorizontalTextAlignment);
            Assert.AreEqual(VerticalTextAlignment.Justify, record.VerticalTextAlignment);
            Assert.AreEqual(true, record.IsTextLocked);
            Assert.AreEqual(TextOrientation.RotRight, record.TextOrientation);


            Assert.AreEqual(49, record.RecordSize);
        }

        [Test]
        public void TestStore()
        {
            TextObjectRecord record = new TextObjectRecord();
            HSSFRichTextString str = new HSSFRichTextString("AB");
            str.ApplyFont(0, 2, (short)0x0018);
            str.ApplyFont(2, 2, (short)0x0320);

            record.HorizontalTextAlignment = HorizontalTextAlignment.Center;
            record.VerticalTextAlignment = VerticalTextAlignment.Justify;
            record.IsTextLocked = (true);
            record.TextOrientation = TextOrientation.RotRight;
            record.Str = (str);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i], "At offset " + i);
        }
    }
}
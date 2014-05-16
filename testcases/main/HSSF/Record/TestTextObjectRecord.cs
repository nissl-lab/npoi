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
    using System.IO;
    using System.Text;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;

    /**
     * Tests that serialization and deserialization of the TextObjectRecord .
     * Test data taken directly from a real Excel file.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestTextObjectRecord
    {

        private static byte[] simpleData = HexRead.ReadFromString(
            "B6 01 12 00 " +
            "12 02 00 00 00 00 00 00" +
            "00 00 0D 00 08 00    00 00" +
            "00 00 " +
            "3C 00 0E 00 " +
            "00 48 65 6C 6C 6F 2C 20 57 6F 72 6C 64 21 " +
            "3C 00 08 " +
            "00 0D 00 00 00 00 00 00 00"
        );

        [Test]
        public void TestRead()
        {

            RecordInputStream is1 = TestcaseRecordInputStream.Create(simpleData);
            TextObjectRecord record = new TextObjectRecord(is1);

            Assert.AreEqual(TextObjectRecord.sid, record.Sid);
            Assert.AreEqual(HorizontalTextAlignment.Left, record.HorizontalTextAlignment);
            Assert.AreEqual(VerticalTextAlignment.Top, record.VerticalTextAlignment);
            Assert.IsTrue(record.IsTextLocked);
            Assert.AreEqual(TextOrientation.None, record.TextOrientation);
            Assert.AreEqual("Hello, World!", record.Str.String);
        }
        [Test]
        public void TestWrite()
        {
            HSSFRichTextString str = new HSSFRichTextString("Hello, World!");

            TextObjectRecord record = new TextObjectRecord();
            record.Str = (/*setter*/str);
            record.HorizontalTextAlignment = (/*setter*/ HorizontalTextAlignment.Left);
            record.VerticalTextAlignment = (/*setter*/ VerticalTextAlignment.Top);
            record.IsTextLocked = (/*setter*/ true);
            record.TextOrientation = (/*setter*/ TextOrientation.None);

            byte[] ser = record.Serialize();
            Assert.AreEqual(ser.Length, simpleData.Length);

            Assert.IsTrue(Arrays.Equals(simpleData, ser));

            //read again
            RecordInputStream is1 = TestcaseRecordInputStream.Create(simpleData);
            record = new TextObjectRecord(is1);
        }

        /**
         * Zero {@link ContinueRecord}s follow a {@link TextObjectRecord} if the text is empty
         */
        [Test]
        public void TestWriteEmpty()
        {
            HSSFRichTextString str = new HSSFRichTextString("");

            TextObjectRecord record = new TextObjectRecord();
            record.Str = (/*setter*/str);

            byte[] ser = record.Serialize();

            int formatDataLen = LittleEndian.GetUShort(ser, 16);
            Assert.AreEqual(0, formatDataLen, "formatDataLength");

            Assert.AreEqual(22, ser.Length); // just the TXO record

            //read again
            RecordInputStream is1 = TestcaseRecordInputStream.Create(ser);
            record = new TextObjectRecord(is1);
            Assert.AreEqual(0, record.Str.Length);
        }

        /**
         * Test that TextObjectRecord Serializes logs records properly.
         */
        [Test]
        public void TestLongRecords()
        {
            int[] length = { 1024, 2048, 4096, 8192, 16384 }; //test against strings of different length
            for (int i = 0; i < length.Length; i++)
            {
                StringBuilder buff = new StringBuilder(length[i]);
                for (int j = 0; j < length[i]; j++)
                {
                    buff.Append("x");
                }
                IRichTextString str = new HSSFRichTextString(buff.ToString());

                TextObjectRecord obj = new TextObjectRecord();
                obj.Str = (/*setter*/str);

                byte[] data = obj.Serialize();
                RecordInputStream is1 = new RecordInputStream(new MemoryStream(data));
                is1.NextRecord();
                TextObjectRecord record = new TextObjectRecord(is1);
                str = record.Str;

                Assert.AreEqual(buff.Length, str.Length);
                Assert.AreEqual(buff.ToString(), str.String);
            }
        }

        /**
         * Test cloning
         */
        [Test]
        public void TestClone()
        {
            String text = "Hello, World";
            HSSFRichTextString str = new HSSFRichTextString(text);

            TextObjectRecord obj = new TextObjectRecord();
            obj.Str = (/*setter*/ str);


            TextObjectRecord Cloned = (TextObjectRecord)obj.Clone();
            Assert.AreEqual(obj.RecordSize, Cloned.RecordSize);
            Assert.AreEqual(obj.HorizontalTextAlignment, Cloned.HorizontalTextAlignment);
            Assert.AreEqual(obj.Str.String, Cloned.Str.String);

            //finally check that the Serialized data is the same
            byte[] src = obj.Serialize();
            byte[] cln = Cloned.Serialize();
            Assert.IsTrue(Arrays.Equals(src, cln));
        }

        /** similar to {@link #simpleData} but with link formula at end of TXO rec*/
        private static byte[] linkData = HexRead.ReadFromString(
                "B6 01 " + // TextObjectRecord.sid
                "1E 00 " + // size 18
                "44 02 02 00 00 00 00 00" +
                "00 00 " +
                "02 00 " + // strLen 2
                "10 00 " + // 16 bytes for 2 format Runs
                "00 00 00 00 " +

                "05 00 " +          // formula size
                "D4 F0 8A 03 " +    // unknownInt
                "24 01 00 13 C0 " + //tRef(T2)
                "13 " +             // ??

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
        public void TestLinkFormula()
        {
            RecordInputStream is1 = new RecordInputStream(new MemoryStream(linkData));
            is1.NextRecord();
            TextObjectRecord rec = new TextObjectRecord(is1);

            Ptg ptg = rec.LinkRefPtg;
            Assert.IsNotNull(ptg);
            Assert.AreEqual(typeof(RefPtg), ptg.GetType());
            RefPtg rptg = (RefPtg)ptg;
            Assert.AreEqual("T2", rptg.ToFormulaString());

            byte[] data2 = rec.Serialize();
            Assert.AreEqual(linkData.Length, data2.Length);
            Assert.IsTrue(Arrays.Equals(linkData, data2));
        }
    }
}

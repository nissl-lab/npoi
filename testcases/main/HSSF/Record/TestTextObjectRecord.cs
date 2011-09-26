
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
    using System.IO;
    using System.Text;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /**
     * Tests that serialization and deserialization of the TextObjectRecord .
     * Test data taken directly from a real Excel file.
     *
     * @author Yegor Kozlov
     */
    [TestClass]
    public class TestTextObjectRecord
    {

        byte[] data = {(byte)0xB6, 0x01, 0x12, 0x00, 0x12, 0x02, 0x00, 0x00, 0x00, 0x00,
                   0x00, 0x00, 0x00, 0x00, 0x0D, 0x00, 0x08, 0x00, 0x00, 0x00, 0x00,
                   0x00, 0x3C, 0x00, 0x1B, 0x00, 0x01, 0x48, 0x00, 0x65, 0x00, 0x6C,
                   0x00, 0x6C, 0x00, 0x6F, 0x00, 0x2C, 0x00, 0x20, 0x00, 0x57, 0x00,
                   0x6F, 0x00, 0x72, 0x00, 0x6C, 0x00, 0x64, 0x00, 0x21, 0x00, 0x3C,
                   0x00, 0x08, 0x00, 0x0D, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

        [TestMethod]
        public void TestRead()
        {


            RecordInputStream is1 = new RecordInputStream(new MemoryStream(data));
            is1.NextRecord();
            TextObjectRecord record = new TextObjectRecord(is1);

            Assert.AreEqual(TextObjectRecord.sid, record.Sid);
            Assert.AreEqual(TextObjectRecord.HORIZONTAL_TEXT_ALIGNMENT_LEFT_ALIGNED, record.HorizontalTextAlignment);
            Assert.AreEqual(TextObjectRecord.VERTICAL_TEXT_ALIGNMENT_TOP, record.VerticalTextAlignment);
            Assert.AreEqual(TextObjectRecord.TEXT_ORIENTATION_NONE, record.TextOrientation);
            Assert.AreEqual("Hello, World!", record.Str.String);

        }
        [TestMethod]
        public void TestWrite()
        {
            HSSFRichTextString str = new HSSFRichTextString("Hello, World!");

            TextObjectRecord record = new TextObjectRecord();
            record.Str = (str);
            record.HorizontalTextAlignment = (TextObjectRecord.HORIZONTAL_TEXT_ALIGNMENT_LEFT_ALIGNED);
            record.VerticalTextAlignment = (TextObjectRecord.VERTICAL_TEXT_ALIGNMENT_TOP);
            record.IsTextLocked = (true);
            record.TextOrientation = (TextObjectRecord.TEXT_ORIENTATION_NONE);

            byte[] ser = record.Serialize();
            //Assert.AreEqual(ser.Length , data.Length);

            //Assert.IsTrue(Arrays.Equals(data, ser));

            //Read again
            RecordInputStream is1 = new RecordInputStream(new MemoryStream(data));
            is1.NextRecord();
            record = new TextObjectRecord(is1);

        }
        /**
         * Zero {@link ContinueRecord}s follow a {@link TextObjectRecord} if the text is empty
         */
        [TestMethod]
        public void TestWriteEmpty()
        {
            HSSFRichTextString str = new HSSFRichTextString("");

            TextObjectRecord record = new TextObjectRecord();
            record.Str = (str);

            byte[] ser = record.Serialize();

            int formatDataLen = NPOI.Util.LittleEndian.GetUShort(ser, 16);
            Assert.AreEqual(0, formatDataLen, "formatDataLength");

            Assert.AreEqual(22, ser.Length); // just the TXO record

            //read again
            RecordInputStream is1 = new RecordInputStream(new MemoryStream(ser));
            is1.NextRecord();
            record = new TextObjectRecord(is1);
            Assert.AreEqual(0, record.Str.Length);
        }
        /**
         * Test that TextObjectRecord Serializes logs records properly.
         */
        [TestMethod]
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
                NPOI.SS.UserModel.IRichTextString str = new HSSFRichTextString(buff.ToString());

                TextObjectRecord obj = new TextObjectRecord();
                obj.Str = (str);

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
        [TestMethod]
        public void TestClone()
        {
            String text = "Hello, World";
            HSSFRichTextString str = new HSSFRichTextString(text);

            TextObjectRecord obj = new TextObjectRecord();
            obj.Str = (str);


            TextObjectRecord cloned = (TextObjectRecord)obj.Clone();
            Assert.AreEqual(obj.RecordSize, cloned.RecordSize);
            Assert.AreEqual(obj.HorizontalTextAlignment, cloned.HorizontalTextAlignment);
            Assert.AreEqual(obj.Str.String, cloned.Str.String);

            //finally check that the serialized data is the same
            byte[] src = obj.Serialize();
            byte[] cln = cloned.Serialize();
            Assert.IsTrue(NPOI.Util.Arrays.Equals(src, cln));
        }
    }
}
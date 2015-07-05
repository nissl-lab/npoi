
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
    using NPOI.HSSF.Record.Cont;
    using System.Text;
    using System.IO;
    using NPOI.Util;

    /**
     * Tests that records size calculates correctly.
     *
     * @author Jason Height (jheight at apache.org)
     */
    [TestFixture]
    public class TestUnicodeString
    {
        private static int MAX_DATA_SIZE = RecordInputStream.MAX_RECORD_DATA_SIZE;
            /** a 4 character string requiring 16 bit encoding */
        private static String STR_16_BIT = "A\u591A\u8A00\u8A9E";

        public TestUnicodeString()
        {

        }

        [Test]
        public void TestSmallStringSize()
        {
            //Test a basic string
            UnicodeString s = MakeUnicodeString("Test");
            ConfirmSize(7, s);

            //Test a small string that is uncompressed
            s = MakeUnicodeString(STR_16_BIT);
            s.OptionFlags = ((byte)0x01);
            ConfirmSize(11, s);

            //Test a compressed small string that has rich text formatting
            s.String = "Test";
            s.OptionFlags = ((byte)0x8);
            UnicodeString.FormatRun r = new UnicodeString.FormatRun((short)0, (short)1);
            s.AddFormatRun(r);
            UnicodeString.FormatRun r2 = new UnicodeString.FormatRun((short)2, (short)2);
            s.AddFormatRun(r2);
            ConfirmSize(17, s);

            //Test a uncompressed small string that has rich text formatting
            s.String = STR_16_BIT;
            s.OptionFlags = ((byte)0x9);
            ConfirmSize(21, s);

            //Test a compressed small string that has rich text and extended text
            s.String = "Test";
            s.OptionFlags = ((byte)0xC);
            ConfirmSize(17, s);
            // Extended phonetics data
            // Minimum size is 14
            // Also adds 4 bytes to hold the length
            s.ExtendedRst=(
                  new UnicodeString.ExtRst()
            );
            ConfirmSize(35, s);

            //Test a uncompressed small string that has rich text and extended text
            s.String = STR_16_BIT;
            s.OptionFlags = ((byte)0xD);
            ConfirmSize(39, s);

            s.ExtendedRst = (null);
            ConfirmSize(21, s);
        }
        [Test]
        public void TestPerfectStringSize()
        {
            //Test a basic string
            UnicodeString s = MakeUnicodeString(MAX_DATA_SIZE - 2 - 1);
            ConfirmSize(MAX_DATA_SIZE, s);

            //Test an uncompressed string
            //Note that we can only ever Get to a maximim size of 8227 since an uncompressed
            //string is1 writing double bytes.
            s = MakeUnicodeString((MAX_DATA_SIZE - 2 - 1) / 2,true);
            s.OptionFlags = (byte)0x1;
            ConfirmSize(MAX_DATA_SIZE - 1, s);
        }
        [Test]
        public void TestPerfectRichStringSize()
        {
            //Test a rich text string
            UnicodeString s = MakeUnicodeString(MAX_DATA_SIZE - 2 - 1 - 8 - 2);
            s.AddFormatRun(new UnicodeString.FormatRun((short)1, (short)0));
            s.AddFormatRun(new UnicodeString.FormatRun((short)2, (short)1));
            s.OptionFlags=((byte)0x8);
            ConfirmSize(MAX_DATA_SIZE, s);

            //Test an uncompressed rich text string
            //Note that we can only ever Get to a maximim size of 8227 since an uncompressed
            //string is1 writing double bytes.
            s = MakeUnicodeString((MAX_DATA_SIZE - 2 - 1 - 8 - 2) / 2,true);
            s.AddFormatRun(new UnicodeString.FormatRun((short)1, (short)0));
            s.AddFormatRun(new UnicodeString.FormatRun((short)2, (short)1));
            s.OptionFlags = ((byte)0x9);
            ConfirmSize(MAX_DATA_SIZE - 1, s);
        }
        [Test]
        public void TestContinuedStringSize()
        {
            UnicodeString s = MakeUnicodeString(MAX_DATA_SIZE - 2 - 1 + 20);
            ConfirmSize(MAX_DATA_SIZE + 4 + 1 + 20, s);
        }

        /** Tests that a string size calculation that fits neatly in two records, the second being a continue*/
        [Test]
        public void TestPerfectContinuedStringSize()
        {
            //Test a basic string
            int strSize = RecordInputStream.MAX_RECORD_DATA_SIZE * 2;
            //String overhead
            strSize -= 3;
            //Continue Record overhead
            strSize -= 4;
            //Continue Record additional byte overhead
            strSize -= 1;
            UnicodeString s = MakeUnicodeString(strSize);
            ConfirmSize(MAX_DATA_SIZE * 2, s);
        }

        [Test]
        public void TestFormatRun()
        {
            UnicodeString.FormatRun fr = new UnicodeString.FormatRun((short)4, (short)0x15c);
            Assert.AreEqual(4, fr.CharacterPos);
            Assert.AreEqual(0x15c, fr.FontIndex);

            MemoryStream baos = new MemoryStream();
            LittleEndianOutputStream out1 = new LittleEndianOutputStream(baos);

            fr.Serialize(out1);

            byte[] b = baos.ToArray();
            Assert.AreEqual(4, b.Length);
            Assert.AreEqual(4, b[0]);
            Assert.AreEqual(0, b[1]);
            Assert.AreEqual(0x5c, b[2]);
            Assert.AreEqual(0x01, b[3]);

            LittleEndianInputStream inp = new LittleEndianInputStream(
                  new MemoryStream(b)
            );
            fr = new UnicodeString.FormatRun(inp);
            Assert.AreEqual(4, fr.CharacterPos);
            Assert.AreEqual(0x15c, fr.FontIndex);
        }

        [Test]
        public void TestExtRstFromEmpty()
        {
            UnicodeString.ExtRst ext = new UnicodeString.ExtRst();

            Assert.AreEqual(0, ext.NumberOfRuns);
            Assert.AreEqual(0, ext.FormattingFontIndex);
            Assert.AreEqual(0, ext.FormattingOptions);
            Assert.AreEqual("", ext.PhoneticText);
            Assert.AreEqual(0, ext.PhRuns.Length);
            Assert.AreEqual(10, ext.DataSize); // Excludes 4 byte header

            MemoryStream baos = new MemoryStream();
            LittleEndianOutputStream out1 = new LittleEndianOutputStream(baos);
            ContinuableRecordOutput cout = new ContinuableRecordOutput(out1, 0xffff);

            ext.Serialize(cout);
            cout.WriteContinue();

            byte[] b = baos.ToArray();
            Assert.AreEqual(20, b.Length);

            // First 4 bytes from the outputstream
            Assert.AreEqual((sbyte)-1, (sbyte)b[0]);
            Assert.AreEqual((sbyte)-1, (sbyte)b[1]);
            Assert.AreEqual(14, b[2]);
            Assert.AreEqual(00, b[3]);

            // Reserved
            Assert.AreEqual(1, b[4]);
            Assert.AreEqual(0, b[5]);
            // Data size
            Assert.AreEqual(10, b[6]);
            Assert.AreEqual(00, b[7]);
            // Font*2
            Assert.AreEqual(0, b[8]);
            Assert.AreEqual(0, b[9]);
            Assert.AreEqual(0, b[10]);
            Assert.AreEqual(0, b[11]);
            // 0 Runs
            Assert.AreEqual(0, b[12]);
            Assert.AreEqual(0, b[13]);
            // Size=0, *2
            Assert.AreEqual(0, b[14]);
            Assert.AreEqual(0, b[15]);
            Assert.AreEqual(0, b[16]);
            Assert.AreEqual(0, b[17]);

            // Last 2 bytes from the outputstream
            Assert.AreEqual(ContinueRecord.sid, b[18]);
            Assert.AreEqual(0, b[19]);


            // Load in again and re-test
            byte[] data = new byte[14];
            Array.Copy(b, 4, data, 0, data.Length);
            LittleEndianInputStream inp = new LittleEndianInputStream(
                  new MemoryStream(data)
            );
            ext = new UnicodeString.ExtRst(inp, data.Length);

            Assert.AreEqual(0, ext.NumberOfRuns);
            Assert.AreEqual(0, ext.FormattingFontIndex);
            Assert.AreEqual(0, ext.FormattingOptions);
            Assert.AreEqual("", ext.PhoneticText);
            Assert.AreEqual(0, ext.PhRuns.Length);
        }

        [Test]
        public void TestExtRstFromData()
        {
            byte[] data = new byte[] {
             01, 00, 0x0C, 00, 
             00, 00, 0x37, 00, 
             00, 00, 
             00, 00, 00, 00, 
             00, 00 // Cruft at the end, as found from real files
       };
            Assert.AreEqual(16, data.Length);

            LittleEndianInputStream inp = new LittleEndianInputStream(
                  new MemoryStream(data)
            );
            UnicodeString.ExtRst ext = new UnicodeString.ExtRst(inp, data.Length);
            Assert.AreEqual(0x0c, ext.DataSize); // Excludes 4 byte header

            Assert.AreEqual(0, ext.NumberOfRuns);
            Assert.AreEqual(0x37, ext.FormattingOptions);
            Assert.AreEqual(0, ext.FormattingFontIndex);
            Assert.AreEqual("", ext.PhoneticText);
            Assert.AreEqual(0, ext.PhRuns.Length);
        }

        [Test]
        public void TestCorruptExtRstDetection()
        {
            byte[] data = new byte[] {
             0x79, 0x79, 0x11, 0x11, 
             0x22, 0x22, 0x33, 0x33, 
       };
            Assert.AreEqual(8, data.Length);

            LittleEndianInputStream inp = new LittleEndianInputStream(
                  new MemoryStream(data)
            );
            UnicodeString.ExtRst ext = new UnicodeString.ExtRst(inp, data.Length);

            // Will be empty
            Assert.AreEqual(ext, new UnicodeString.ExtRst());

            // If written, will be the usual size
            Assert.AreEqual(10, ext.DataSize); // Excludes 4 byte header

            // Is empty
            Assert.AreEqual(0, ext.NumberOfRuns);
            Assert.AreEqual(0, ext.FormattingOptions);
            Assert.AreEqual(0, ext.FormattingFontIndex);
            Assert.AreEqual("", ext.PhoneticText);
            Assert.AreEqual(0, ext.PhRuns.Length);
        }
        [Test]
        public void TestExtRstEqualsAndHashCode()
        {
            byte[] buf = new byte[200];
            LittleEndianByteArrayOutputStream bos = new LittleEndianByteArrayOutputStream(buf, 0);
            String str = "\u1d02\u1d12\u1d22";
            bos.WriteShort(1);
            bos.WriteShort(5 * LittleEndianConsts.SHORT_SIZE + str.Length * 2 + 3 * LittleEndianConsts.SHORT_SIZE + 2); // data size
            bos.WriteShort(0x4711);
            bos.WriteShort(0x0815);
            bos.WriteShort(1);
            bos.WriteShort(str.Length);
            bos.WriteShort(str.Length);
            StringUtil.PutUnicodeLE(str, bos);
            bos.WriteShort(1);
            bos.WriteShort(1);
            bos.WriteShort(3);
            bos.WriteShort(42);

            LittleEndianByteArrayInputStream in1 = new LittleEndianByteArrayInputStream(buf, 0, bos.WriteIndex);
            UnicodeString.ExtRst extRst1 = new UnicodeString.ExtRst(in1, bos.WriteIndex);
            in1 = new LittleEndianByteArrayInputStream(buf, 0, bos.WriteIndex);
            UnicodeString.ExtRst extRst2 = new UnicodeString.ExtRst(in1, bos.WriteIndex);

            Assert.AreEqual(extRst1, extRst2);
            Assert.AreEqual(extRst1.GetHashCode(), extRst2.GetHashCode());
        }

        private static void ConfirmSize(int expectedSize, UnicodeString s)
        {
            ConfirmSize(expectedSize, s, 0);
        }
        /**
         * Note - a value of zero for <c>amountUsedInCurrentRecord</c> would only ever occur just
         * after a {@link ContinueRecord} had been started.  In the initial {@link SSTRecord} this 
         * value starts at 8 (for the first {@link UnicodeString} written).  In general, it can be
         * any value between 0 and {@link #MAX_DATA_SIZE}
         */
        private static void ConfirmSize(int expectedSize, UnicodeString s, int amountUsedInCurrentRecord)
        {
            ContinuableRecordOutput out1 = ContinuableRecordOutput.CreateForCountingOnly();
            out1.WriteContinue();
            for (int i = amountUsedInCurrentRecord; i > 0; i--)
            {
                out1.WriteByte(0);
            }
            int size0 = out1.TotalSize;
            s.Serialize(out1);
            int size1 = out1.TotalSize;
            int actualSize = size1 - size0;
            Assert.AreEqual(expectedSize, actualSize);
        }

        private static UnicodeString MakeUnicodeString(String s)
        {
            UnicodeString st = new UnicodeString(s);
            st.OptionFlags = ((byte)0);
            return st;
        }

        private static UnicodeString MakeUnicodeString(int numChars)
        {
            StringBuilder b = new StringBuilder(numChars);
            for (int i = 0; i < numChars; i++)
            {
                b.Append(i % 10);
            }
            return MakeUnicodeString(b.ToString());
        }
        /**
         * @param is16Bit if <c>true</c> the created string will have characters > 0x00FF
         * @return a string of the specified number of characters
         */
        private static UnicodeString MakeUnicodeString(int numChars, bool is16Bit)
        {
            StringBuilder b = new StringBuilder(numChars);
            int charBase = is16Bit ? 0x8A00 : 'A';
            for (int i = 0; i < numChars; i++)
            {
                char ch = (char)((i % 16) + charBase);
                b.Append(ch);
            }
            return MakeUnicodeString(b.ToString());
        }
    }
}

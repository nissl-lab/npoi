
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

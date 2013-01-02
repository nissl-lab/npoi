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

namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using System.IO;
    using TestCases.HSSF;
    using NUnit.Framework;

    using NPOI.SS.UserModel;

    [TestFixture]
    public class TestHSSFRichTextString
    {
        [Test]
        public void TestApplyFont()
        {

            HSSFRichTextString r = new HSSFRichTextString("Testing");
            Assert.AreEqual(0, r.NumFormattingRuns);
            r.ApplyFont(2, 4, new HSSFFont((short)1, null));
            Assert.AreEqual(2, r.NumFormattingRuns);
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(0));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(1));
            Assert.AreEqual(1, r.GetFontAtIndex(2));
            Assert.AreEqual(1, r.GetFontAtIndex(3));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(4));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(5));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(6));

            r.ApplyFont(6, 7, new HSSFFont((short)2, null));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(0));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(1));
            Assert.AreEqual(1, r.GetFontAtIndex(2));
            Assert.AreEqual(1, r.GetFontAtIndex(3));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(4));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(5));
            Assert.AreEqual(2, r.GetFontAtIndex(6));

            r.ApplyFont(HSSFRichTextString.NO_FONT);
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(0));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(1));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(2));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(3));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(4));
            Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(5));

            r.ApplyFont(new HSSFFont((short)1, null));
            Assert.AreEqual(1, r.GetFontAtIndex(0));
            Assert.AreEqual(1, r.GetFontAtIndex(1));
            Assert.AreEqual(1, r.GetFontAtIndex(2));
            Assert.AreEqual(1, r.GetFontAtIndex(3));
            Assert.AreEqual(1, r.GetFontAtIndex(4));
            Assert.AreEqual(1, r.GetFontAtIndex(5));
            Assert.AreEqual(1, r.GetFontAtIndex(6));

        }
        [Test]
        public void TestClearFormatting()
        {

            HSSFRichTextString r = new HSSFRichTextString("Testing");
            Assert.AreEqual(0, r.NumFormattingRuns);
            r.ApplyFont(2, 4, new HSSFFont((short)1, null));
            Assert.AreEqual(2, r.NumFormattingRuns);
            r.ClearFormatting();
            Assert.AreEqual(0, r.NumFormattingRuns);
        }
        /**
  * Test case proposed in Bug 40520:  formated twice => will format whole String
  */
        [Test]
        public void Test40520_1()
        {

            short font = 3;

            HSSFRichTextString r = new HSSFRichTextString("f0_123456789012345678901234567890123456789012345678901234567890");

            r.ApplyFont(0, 7, font);
            r.ApplyFont(5, 9, font);

            for (int i = 0; i < 7; i++) Assert.AreEqual(font, r.GetFontAtIndex(i));
            for (int i = 5; i < 9; i++) Assert.AreEqual(font, r.GetFontAtIndex(i));
            for (int i = 9; i < r.Length; i++) Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(i));
        }

        /**
         * Test case proposed in Bug 40520:  overlapped range => will format whole String
         */
        [Test]
        public void Test40520_2()
        {

            short font = 3;
            HSSFRichTextString r = new HSSFRichTextString("f0_123456789012345678901234567890123456789012345678901234567890");

            r.ApplyFont(0, 2, font);
            for (int i = 0; i < 2; i++) Assert.AreEqual(font, r.GetFontAtIndex(i));
            for (int i = 2; i < r.Length; i++) Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(i));

            r.ApplyFont(0, 2, font);
            for (int i = 0; i < 2; i++) Assert.AreEqual(font, r.GetFontAtIndex(i));
            for (int i = 2; i < r.Length; i++) Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(i));
        }

        /**
         * Test case proposed in Bug 40520:  formated twice => will format whole String
         */
        [Test]
        public void Test40520_3()
        {

            short font = 3;
            HSSFRichTextString r = new HSSFRichTextString("f0_123456789012345678901234567890123456789012345678901234567890");

            // wrong order => will format 0-6
            r.ApplyFont(0, 2, font);
            r.ApplyFont(5, 7, font);
            r.ApplyFont(0, 2, font);

            r.ApplyFont(0, 2, font);
            for (int i = 0; i < 2; i++) Assert.AreEqual(font, r.GetFontAtIndex(i));
            for (int i = 2; i < 5; i++) Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(i));
            for (int i = 5; i < 7; i++) Assert.AreEqual(font, r.GetFontAtIndex(i));
            for (int i = 7; i < r.Length; i++) Assert.AreEqual(HSSFRichTextString.NO_FONT, r.GetFontAtIndex(i));
        }
    }
}
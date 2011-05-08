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
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.SS.UserModel;

    [TestClass]
    public class TestHSSFRichTextString
    {
        [TestMethod]
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
        [TestMethod]
        public void TestClearFormatting()
        {

            HSSFRichTextString r = new HSSFRichTextString("Testing");
            Assert.AreEqual(0, r.NumFormattingRuns);
            r.ApplyFont(2, 4, new HSSFFont((short)1, null));
            Assert.AreEqual(2, r.NumFormattingRuns);
            r.ClearFormatting();
            Assert.AreEqual(0, r.NumFormattingRuns);
        }
    }
}
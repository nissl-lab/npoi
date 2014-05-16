/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using System.IO;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Record;

    /**
     * Test <c>HSSFTextbox</c>.
     *
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestFixture]
    public class TestHSSFTextbox
    {

        /**
         * Test that accessors to horizontal and vertical alignment work properly
         */
        [Test]
        public void TestAlignment()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sh1 = wb.CreateSheet();
            HSSFPatriarch patriarch = sh1.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor(0, 0, 0, 0, 1, 1, 6, 4)) as HSSFTextbox;
            HSSFRichTextString str = new HSSFRichTextString("Hello, World");
            textbox.String = (str);
            textbox.HorizontalAlignment = HorizontalTextAlignment.Center;
            textbox.VerticalAlignment = VerticalTextAlignment.Center;

            Assert.AreEqual(HorizontalTextAlignment.Center, textbox.HorizontalAlignment);
            Assert.AreEqual(VerticalTextAlignment.Center, textbox.VerticalAlignment);
        }

        /**
         * Excel requires at least one format run in HSSFTextbox.
         * When inserting text make sure that if font is not set we must set the default one.
         */
        [Test]
        public void TestSetDeafultTextFormat()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFTextbox textbox1 = patriarch.CreateTextbox(new HSSFClientAnchor(0, 0, 0, 0, 1, 1, 3, 3)) as HSSFTextbox;
            HSSFRichTextString rt1 = new HSSFRichTextString("Hello, World!");
            Assert.AreEqual(0, rt1.NumFormattingRuns);
            textbox1.String=(rt1);

            HSSFRichTextString rt2 = (HSSFRichTextString)textbox1.String;
            Assert.AreEqual(1, rt2.NumFormattingRuns);
            Assert.AreEqual(HSSFRichTextString.NO_FONT, rt2.GetFontOfFormattingRun(0));
        }

    }
}

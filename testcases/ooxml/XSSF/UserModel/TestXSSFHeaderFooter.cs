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

using System;
using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel
{
    /**
     * Tests for {@link XSSFHeaderFooter}
     */
    [TestFixture]
    public class TestXSSFHeaderFooter
    {
        [Test]
        public void TestStripFields()
        {
            String simple = "I am a Test header";
            String withPage = "I am a&P Test header";
            String withLots = "I&A am&N a&P Test&T header&U";
            String withFont = "I&22 am a&\"Arial,bold\" Test header";
            String withOtherAnds = "I am a&P Test header&&";
            String withOtherAnds2 = "I am a&P Test header&a&b";

            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(simple));
            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(withPage));
            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(withLots));
            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(withFont));
            ClassicAssert.AreEqual(simple + "&&", XSSFOddHeader.StripFields(withOtherAnds));
            ClassicAssert.AreEqual(simple + "&a&b", XSSFOddHeader.StripFields(withOtherAnds2));

            // Now Test the default strip flag
            XSSFEvenHeader head = new XSSFEvenHeader(new CT_HeaderFooter());
            head.Center = ("Center");
            head.Left = ("In the left");

            ClassicAssert.AreEqual("In the left", head.Left);
            ClassicAssert.AreEqual("Center", head.Center);
            ClassicAssert.AreEqual("", head.Right);

            head.Left = ("Top &P&F&D Left");
            ClassicAssert.AreEqual("Top &P&F&D Left", head.Left);
            ClassicAssert.IsFalse(head.AreFieldsStripped());

            head.SetAreFieldsStripped(true);
            ClassicAssert.AreEqual("Top  Left", head.Left);
            ClassicAssert.IsTrue(head.AreFieldsStripped());

            // Now even more complex
            head.Center = ("HEADER TEXT &P&N&D&T&Z&F&F&A&V");
            ClassicAssert.AreEqual("HEADER TEXT &V", head.Center);
        }
        [Test]
        public void TestGetSetCenterLeftRight()
        {

            XSSFOddFooter footer = new XSSFOddFooter(new CT_HeaderFooter());
            ClassicAssert.AreEqual("", footer.Center);
            footer.Center = ("My first center section");
            ClassicAssert.AreEqual("My first center section", footer.Center);
            footer.Center = ("No, let's update the center section");
            ClassicAssert.AreEqual("No, let's update the center section", footer.Center);
            footer.Left = ("And add a left one");
            footer.Right = ("Finally the right section is Added");
            ClassicAssert.AreEqual("And add a left one", footer.Left);
            ClassicAssert.AreEqual("Finally the right section is Added", footer.Right);

            // Test changing the three sections value
            footer.Center = ("Second center version");
            footer.Left = ("Second left version");
            footer.Right = ("Second right version");
            ClassicAssert.AreEqual("Second center version", footer.Center);
            ClassicAssert.AreEqual("Second left version", footer.Left);
            ClassicAssert.AreEqual("Second right version", footer.Right);

        }

        // TODO Rest of Tests
    }


}


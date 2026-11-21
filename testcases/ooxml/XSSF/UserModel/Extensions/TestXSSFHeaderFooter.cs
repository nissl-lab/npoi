/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XSSF.UserModel.Extensions
{
    using NPOI.XSSF.UserModel;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using NPOI.XSSF.UserModel.Extensions;
    [TestFixture]
    public class TestXSSFHeaderFooter
    {

        private XSSFWorkbook wb;
        private XSSFSheet sheet;
        private XSSFHeaderFooter hO;
        private XSSFHeaderFooter hE;
        private XSSFHeaderFooter hF;
        private XSSFHeaderFooter fO;
        private XSSFHeaderFooter fE;
        private XSSFHeaderFooter fF;

        [SetUp]
        public void before()
        {
            wb = new XSSFWorkbook();
            sheet = wb.CreateSheet() as XSSFSheet;
            hO = (XSSFHeaderFooter) sheet.OddHeader;
            hE = (XSSFHeaderFooter) sheet.EvenHeader;
            hF = (XSSFHeaderFooter) sheet.FirstHeader;
            fO = (XSSFHeaderFooter) sheet.OddFooter;
            fE = (XSSFHeaderFooter) sheet.EvenFooter;
            fF = (XSSFHeaderFooter) sheet.FirstFooter;
        }

        [TearDown]
        public void After()
        {
            wb.Close();
        }

        [Test]
        public void TestGetHeaderFooter()
        {
            CT_HeaderFooter ctHf;
            ctHf = hO.GetHeaderFooter();
            ClassicAssert.IsNotNull(ctHf);
            ctHf = hE.GetHeaderFooter();
            ClassicAssert.IsNotNull(ctHf);
            ctHf = hF.GetHeaderFooter();
            ClassicAssert.IsNotNull(ctHf);
            ctHf = fO.GetHeaderFooter();
            ClassicAssert.IsNotNull(ctHf);
            ctHf = fE.GetHeaderFooter();
            ClassicAssert.IsNotNull(ctHf);
            ctHf = fF.GetHeaderFooter();
            ClassicAssert.IsNotNull(ctHf);
        }

        [Test]
        public void TestGetValue()
        {
            ClassicAssert.AreEqual("", hO.Value);
            ClassicAssert.AreEqual("", hE.Value);
            ClassicAssert.AreEqual("", hF.Value);
            ClassicAssert.AreEqual("", fO.Value);
            ClassicAssert.AreEqual("", fE.Value);
            ClassicAssert.AreEqual("", fF.Value);
            hO.Left = "Left value";
            hO.Center = "Center value";
            hO.Right = "Right value";
            hE.Left = "LeftEvalue";
            hE.Center = "CenterEvalue";
            hE.Right = "RightEvalue";
            hF.Left = "LeftFvalue";
            hF.Center = "CenterFvalue";
            hF.Right = "RightFvalue";
            ClassicAssert.AreEqual("&CCenter value&LLeft value&RRight value", hO.Value);
            ClassicAssert.AreEqual("&CCenterEvalue&LLeftEvalue&RRightEvalue", hE.Value);
            ClassicAssert.AreEqual("&CCenterFvalue&LLeftFvalue&RRightFvalue", hF.Value);
            fO.Left = "Left value1";
            fO.Center = "Center value1";
            fO.Right = "Right value1";
            fE.Left = "LeftEvalue1";
            fE.Center = "CenterEvalue1";
            fE.Right = "RightEvalue1";
            fF.Left = "LeftFvalue1";
            fF.Center = "CenterFvalue1";
            fF.Right = "RightFvalue1";
            ClassicAssert.AreEqual("&CCenter value1&LLeft value1&RRight value1", fO.Value);
            ClassicAssert.AreEqual("&CCenterEvalue1&LLeftEvalue1&RRightEvalue1", fE.Value);
            ClassicAssert.AreEqual("&CCenterFvalue1&LLeftFvalue1&RRightFvalue1", fF.Value);
        }

        [Ignore("Test not yet created")]
        public void TestAreFieldsStripped()
        {
            ClassicAssert.Fail("Not yet implemented");
        }

        [Ignore("Test not yet created")]
        public void TestSetAreFieldsStripped()
        {
            ClassicAssert.Fail("Not yet implemented");
        }

        [Test]
        public void TestStripFields()
        {
            string simple = "I am a test header";
            string withPage = "I am a&P test header";
            string withLots = "I&A am&N a&P test&T header&U";
            string withFont = "I&22 am a&\"Arial,bold\" test header";
            string withOtherAnds = "I am a&P test header&&";
            string withOtherAnds2 = "I am a&P test header&a&b";

            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(simple));
            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(withPage));
            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(withLots));
            ClassicAssert.AreEqual(simple, XSSFOddHeader.StripFields(withFont));
            ClassicAssert.AreEqual(simple + "&&", XSSFOddHeader.StripFields(withOtherAnds));
            ClassicAssert.AreEqual(simple + "&a&b", XSSFOddHeader.StripFields(withOtherAnds2));

            // Now test the default strip flag
            hE.Center = "Center";
            hE.Left = "In the left";

            ClassicAssert.AreEqual("In the left", hE.Left);
            ClassicAssert.AreEqual("Center", hE.Center);
            ClassicAssert.AreEqual("", hE.Right);

            hE.Left = "Top &P&F&D Left";
            ClassicAssert.AreEqual("Top &P&F&D Left", hE.Left);
            ClassicAssert.IsFalse(hE.AreFieldsStripped());

            hE.SetAreFieldsStripped(true);
            ClassicAssert.AreEqual("Top  Left", hE.Left);
            ClassicAssert.IsTrue(hE.AreFieldsStripped());

            // Now even more complex
            hE.Center = "HEADER TEXT &P&N&D&T&Z&F&F&A&V";
            ClassicAssert.AreEqual("HEADER TEXT &V", hE.Center);
        }

        [Test]
        public void TestGetCenter()
        {
            ClassicAssert.AreEqual("", hO.Center);
            ClassicAssert.AreEqual("", hE.Center);
            ClassicAssert.AreEqual("", hF.Center);
            ClassicAssert.AreEqual("", fO.Center);
            ClassicAssert.AreEqual("", fE.Center);
            ClassicAssert.AreEqual("", fF.Center);
            hO.Center = "Center value";
            hE.Center = "CenterEvalue";
            hF.Center = "CenterFvalue";
            ClassicAssert.AreEqual("Center value", hO.Center);
            ClassicAssert.AreEqual("CenterEvalue", hE.Center);
            ClassicAssert.AreEqual("CenterFvalue", hF.Center);
            fO.Center = "Center value1";
            fE.Center = "CenterEvalue1";
            fF.Center = "CenterFvalue1";
            ClassicAssert.AreEqual("Center value1", fO.Center);
            ClassicAssert.AreEqual("CenterEvalue1", fE.Center);
            ClassicAssert.AreEqual("CenterFvalue1", fF.Center);
        }

        [Test]
        public void TestGetLeft()
        {
            ClassicAssert.AreEqual("", hO.Left);
            ClassicAssert.AreEqual("", hE.Left);
            ClassicAssert.AreEqual("", hF.Left);
            ClassicAssert.AreEqual("", fO.Left);
            ClassicAssert.AreEqual("", fE.Left);
            ClassicAssert.AreEqual("", fF.Left);
            hO.Left = "Left value";
            hE.Left = "LeftEvalue";
            hF.Left = "LeftFvalue";
            ClassicAssert.AreEqual("Left value", hO.Left);
            ClassicAssert.AreEqual("LeftEvalue", hE.Left);
            ClassicAssert.AreEqual("LeftFvalue", hF.Left);
            fO.Left = "Left value1";
            fE.Left = "LeftEvalue1";
            fF.Left = "LeftFvalue1";
            ClassicAssert.AreEqual("Left value1", fO.Left);
            ClassicAssert.AreEqual("LeftEvalue1", fE.Left);
            ClassicAssert.AreEqual("LeftFvalue1", fF.Left);
        }

        [Test]
        public void TestGetRight()
        {
            ClassicAssert.AreEqual("", hO.Value);
            ClassicAssert.AreEqual("", hE.Value);
            ClassicAssert.AreEqual("", hF.Value);
            ClassicAssert.AreEqual("", fO.Value);
            ClassicAssert.AreEqual("", fE.Value);
            ClassicAssert.AreEqual("", fF.Value);
            hO.Right = "Right value";
            hE.Right = "RightEvalue";
            hF.Right = "RightFvalue";
            ClassicAssert.AreEqual("Right value", hO.Right);
            ClassicAssert.AreEqual("RightEvalue", hE.Right);
            ClassicAssert.AreEqual("RightFvalue", hF.Right);
            fO.Right = "Right value1";
            fE.Right = "RightEvalue1";
            fF.Right = "RightFvalue1";
            ClassicAssert.AreEqual("Right value1", fO.Right);
            ClassicAssert.AreEqual("RightEvalue1", fE.Right);
            ClassicAssert.AreEqual("RightFvalue1", fF.Right);
        }

        [Test]
        public void TestSetCenter()
        {
            ClassicAssert.AreEqual("", hO.Value);
            ClassicAssert.AreEqual("", hE.Value);
            ClassicAssert.AreEqual("", hF.Value);
            ClassicAssert.AreEqual("", fO.Value);
            ClassicAssert.AreEqual("", fE.Value);
            ClassicAssert.AreEqual("", fF.Value);
            hO.Center = "Center value";
            hE.Center = "CenterEvalue";
            hF.Center = "CenterFvalue";
            ClassicAssert.AreEqual("&CCenter value", hO.Value);
            ClassicAssert.AreEqual("&CCenterEvalue", hE.Value);
            ClassicAssert.AreEqual("&CCenterFvalue", hF.Value);
            fO.Center = "Center value1";
            fE.Center = "CenterEvalue1";
            fF.Center = "CenterFvalue1";
            ClassicAssert.AreEqual("&CCenter value1", fO.Value);
            ClassicAssert.AreEqual("&CCenterEvalue1", fE.Value);
            ClassicAssert.AreEqual("&CCenterFvalue1", fF.Value);
        }

        [Test]
        public void TestSetLeft()
        {
            ClassicAssert.AreEqual("", hO.Value);
            ClassicAssert.AreEqual("", hE.Value);
            ClassicAssert.AreEqual("", hF.Value);
            ClassicAssert.AreEqual("", fO.Value);
            ClassicAssert.AreEqual("", fE.Value);
            ClassicAssert.AreEqual("", fF.Value);
            hO.Left = "Left value";
            hE.Left = "LeftEvalue";
            hF.Left = "LeftFvalue";
            ClassicAssert.AreEqual("&LLeft value", hO.Value);
            ClassicAssert.AreEqual("&LLeftEvalue", hE.Value);
            ClassicAssert.AreEqual("&LLeftFvalue", hF.Value);
            fO.Left = "Left value1";
            fE.Left = "LeftEvalue1";
            fF.Left = "LeftFvalue1";
            ClassicAssert.AreEqual("&LLeft value1", fO.Value);
            ClassicAssert.AreEqual("&LLeftEvalue1", fE.Value);
            ClassicAssert.AreEqual("&LLeftFvalue1", fF.Value);
        }

        [Test]
        public void TestSetRight()
        {
            ClassicAssert.AreEqual("", hO.Value);
            ClassicAssert.AreEqual("", hE.Value);
            ClassicAssert.AreEqual("", hF.Value);
            ClassicAssert.AreEqual("", fO.Value);
            ClassicAssert.AreEqual("", fE.Value);
            ClassicAssert.AreEqual("", fF.Value);
            hO.Right = "Right value";
            hE.Right = "RightEvalue";
            hF.Right = "RightFvalue";
            ClassicAssert.AreEqual("&RRight value", hO.Value);
            ClassicAssert.AreEqual("&RRightEvalue", hE.Value);
            ClassicAssert.AreEqual("&RRightFvalue", hF.Value);
            fO.Right = "Right value1";
            fE.Right = "RightEvalue1";
            fF.Right = "RightFvalue1";
            ClassicAssert.AreEqual("&RRight value1", fO.Value);
            ClassicAssert.AreEqual("&RRightEvalue1", fE.Value);
            ClassicAssert.AreEqual("&RRightFvalue1", fF.Value);
        }



        [Test]
        public void TestGetSetCenterLeftRight()
        {
            ClassicAssert.AreEqual("", fO.Center);
            fO.Center = "My first center section";
            ClassicAssert.AreEqual("My first center section", fO.Center);
            fO.Center = "No, let's update the center section";
            ClassicAssert.AreEqual("No, let's update the center section", fO.Center);
            fO.Left = "And add a left one";
            fO.Right = "Finally the right section is added";
            ClassicAssert.AreEqual("And add a left one", fO.Left);
            ClassicAssert.AreEqual("Finally the right section is added", fO.Right);

            // Test changing the three sections value
            fO.Center = "Second center version";
            fO.Left = "Second left version";
            fO.Right = "Second right version";
            ClassicAssert.AreEqual("Second center version", fO.Center);
            ClassicAssert.AreEqual("Second left version", fO.Left);
            ClassicAssert.AreEqual("Second right version", fO.Right);

        }
    }
}

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

using NPOI.HWPF;

using NPOI.HWPF.UserModel;
using NUnit.Framework;
namespace TestCases.HWPF.UserModel
{
    /**
     * Tests for the handling of header stories into headers, footers etc
     */
    [TestFixture]
    public class TestHeaderStories
    {
        private HWPFDocument none;
        private HWPFDocument header;
        private HWPFDocument footer;
        private HWPFDocument headerFooter;
        private HWPFDocument oddEven;
        private HWPFDocument diffFirst;
        private HWPFDocument unicode;
        private HWPFDocument withFields;
        [SetUp]
        public void SetUp()
        {

            none = HWPFTestDataSamples.OpenSampleFile("NoHeadFoot.doc");
            header = HWPFTestDataSamples.OpenSampleFile("ThreeColHead.doc");
            footer = HWPFTestDataSamples.OpenSampleFile("ThreeColFoot.doc");
            headerFooter = HWPFTestDataSamples.OpenSampleFile("SimpleHeadThreeColFoot.doc");
            oddEven = HWPFTestDataSamples.OpenSampleFile("PageSpecificHeadFoot.doc");
            diffFirst = HWPFTestDataSamples.OpenSampleFile("DiffFirstPageHeadFoot.doc");
            unicode = HWPFTestDataSamples.OpenSampleFile("HeaderFooterUnicode.doc");
            withFields = HWPFTestDataSamples.OpenSampleFile("HeaderWithMacros.doc");
        }
        [Test]
        public void TestNone()
        {
            HeaderStories hs = new HeaderStories(none);

            Assert.IsNull(hs.GetPlcfHdd());
            Assert.AreEqual(0, hs.GetRange().Text.Length);
        }
        [Test]
        public void TestHeader()
        {
            HeaderStories hs = new HeaderStories(header);

            Assert.AreEqual(60, hs.GetRange().Text.Length);

            // Should have the usual 6 Separaters
            // Then all 6 of the different header/footer kinds
            // Finally a terminater
            Assert.AreEqual(13, hs.GetPlcfHdd().Length);

            Assert.AreEqual(215, hs.GetRange().StartOffset);

            Assert.AreEqual(0, hs.GetPlcfHdd().GetProperty(0).Start);
            Assert.AreEqual(3, hs.GetPlcfHdd().GetProperty(1).Start);
            Assert.AreEqual(6, hs.GetPlcfHdd().GetProperty(2).Start);
            Assert.AreEqual(6, hs.GetPlcfHdd().GetProperty(3).Start);
            Assert.AreEqual(9, hs.GetPlcfHdd().GetProperty(4).Start);
            Assert.AreEqual(12, hs.GetPlcfHdd().GetProperty(5).Start);

            Assert.AreEqual(12, hs.GetPlcfHdd().GetProperty(6).Start);
            Assert.AreEqual(12, hs.GetPlcfHdd().GetProperty(7).Start);
            Assert.AreEqual(59, hs.GetPlcfHdd().GetProperty(8).Start);
            Assert.AreEqual(59, hs.GetPlcfHdd().GetProperty(9).Start);
            Assert.AreEqual(59, hs.GetPlcfHdd().GetProperty(10).Start);
            Assert.AreEqual(59, hs.GetPlcfHdd().GetProperty(11).Start);

            Assert.AreEqual(59, hs.GetPlcfHdd().GetProperty(12).Start);

            Assert.AreEqual("\u0003\r\r", hs.FootnoteSeparator);
            Assert.AreEqual("\u0004\r\r", hs.FootnoteContSeparator);
            Assert.AreEqual("", hs.FootnoteContNote);
            Assert.AreEqual("\u0003\r\r", hs.EndnoteSeparator);
            Assert.AreEqual("\u0004\r\r", hs.EndnoteContSeparator);
            Assert.AreEqual("", hs.EndnoteContNote);

            Assert.AreEqual("", hs.FirstHeader);
            Assert.AreEqual("", hs.EvenHeader);
            Assert.AreEqual("First header column!\tMid header Right header!\r\r", hs.OddHeader);

            Assert.AreEqual("", hs.FirstFooter);
            Assert.AreEqual("", hs.EvenFooter);
            Assert.AreEqual("", hs.OddFooter);
        }
        [Test]
        public void TestFooter()
        {
            HeaderStories hs = new HeaderStories(footer);

            Assert.AreEqual("", hs.FirstHeader);
            Assert.AreEqual("", hs.EvenHeader);
            Assert.AreEqual("", hs.OddHeader); // Was \r\r but Gets emptied

            Assert.AreEqual("", hs.FirstFooter);
            Assert.AreEqual("", hs.EvenFooter);
            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.OddFooter);
        }
        [Test]
        public void TestHeaderFooter()
        {
            HeaderStories hs = new HeaderStories(headerFooter);

            Assert.AreEqual("", hs.FirstHeader);
            Assert.AreEqual("", hs.EvenHeader);
            Assert.AreEqual("I am some simple header text here\r\r\r", hs.OddHeader);

            Assert.AreEqual("", hs.FirstFooter);
            Assert.AreEqual("", hs.EvenFooter);
            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.OddFooter);
        }
        [Test]
        public void TestOddEven()
        {
            HeaderStories hs = new HeaderStories(oddEven);

            Assert.AreEqual("", hs.FirstHeader);
            Assert.AreEqual("[This is an Even Page, with a Header]\u0007August 20, 2008\u0007\u0007\r\r",
                    hs.EvenHeader);
            Assert.AreEqual("August 20, 2008\u0007[ODD Page Header text]\u0007\u0007\r\r", hs
                    .OddHeader);

            Assert.AreEqual("", hs.FirstFooter);
            Assert.AreEqual(
                    "\u0007Page \u0013 PAGE  \\* MERGEFORMAT \u00142\u0015\u0007\u0007\u0007\u0007\u0007\u0007\u0007This is a simple footer on the second page\r\r",
                    hs.EvenFooter);
            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.OddFooter);

            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.GetFooter(1));
            Assert.AreEqual(
                    "\u0007Page \u0013 PAGE  \\* MERGEFORMAT \u00142\u0015\u0007\u0007\u0007\u0007\u0007\u0007\u0007This is a simple footer on the second page\r\r",
                    hs.GetFooter(2));
            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.GetFooter(3));
        }
        [Test]
        public void TestFirst()
        {
            HeaderStories hs = new HeaderStories(diffFirst);

            Assert.AreEqual("I am the header on the first page, and I\u2019m nice and simple\r\r", hs
                    .FirstHeader);
            Assert.AreEqual("", hs.EvenHeader);
            Assert.AreEqual("First header column!\tMid header Right header!\r\r", hs.OddHeader);

            Assert.AreEqual("The footer of the first page\r\r", hs.FirstFooter);
            Assert.AreEqual("", hs.EvenFooter);
            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.OddFooter);

            Assert.AreEqual("The footer of the first page\r\r", hs.GetFooter(1));
            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.GetFooter(2));
            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\r\r", hs.GetFooter(3));
        }
        [Test]
        public void TestUnicode()
        {
            HeaderStories hs = new HeaderStories(unicode);

            Assert.AreEqual("", hs.FirstHeader);
            Assert.AreEqual("", hs.EvenHeader);
            Assert.AreEqual("This is a simple header, with a \u20ac euro symbol in it.\r\r\r", hs
                    .OddHeader);

            Assert.AreEqual("", hs.FirstFooter);
            Assert.AreEqual("", hs.EvenFooter);
            Assert.AreEqual("The footer, with Moli\u00e8re, has Unicode in it.\r\r", hs.OddFooter);
        }
        [Test]
        public void TestWithFields()
        {
            HeaderStories hs = new HeaderStories(withFields);
            Assert.IsFalse(hs.AreFieldsStripped);

            Assert.AreEqual(
                    "HEADER GOES HERE. 8/12/2008 \u0013 AUTHOR   \\* MERGEFORMAT \u0014Eric Roch\u0015\r\r\r",
                    hs.OddHeader);

            // Now turn on stripping
            hs.AreFieldsStripped = true;
            Assert.AreEqual("HEADER GOES HERE. 8/12/2008 Eric Roch\r\r\r", hs.OddHeader);
        }
    }
}

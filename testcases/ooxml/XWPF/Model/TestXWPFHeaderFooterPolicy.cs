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

namespace TestCases.XWPF.Model
{
    using NPOI.XWPF.Model;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    /**
     * Tests for XWPF Header Footer Stuff
     */
    [TestFixture]
    public class TestXWPFHeaderFooterPolicy
    {
        private XWPFDocument noHeader;
        private XWPFDocument header;
        private XWPFDocument headerFooter;
        private XWPFDocument footer;
        private XWPFDocument oddEven;
        private XWPFDocument diffFirst;
        [SetUp]
        public void SetUp()
        {

            noHeader = XWPFTestDataSamples.OpenSampleDocument("NoHeadFoot.docx");
            header = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");
            headerFooter = XWPFTestDataSamples.OpenSampleDocument("SimpleHeadThreeColFoot.docx");
            footer = XWPFTestDataSamples.OpenSampleDocument("FancyFoot.docx");
            oddEven = XWPFTestDataSamples.OpenSampleDocument("PageSpecificHeadFoot.docx");
            diffFirst = XWPFTestDataSamples.OpenSampleDocument("DiffFirstPageHeadFoot.docx");
        }

        [Test]
        public void TestPolicy()
        {
            XWPFHeaderFooterPolicy policy;

            policy = noHeader.GetHeaderFooterPolicy();
            ClassicAssert.IsNull(policy.GetDefaultHeader());
            ClassicAssert.IsNull(policy.GetDefaultFooter());

            ClassicAssert.IsNull(policy.GetHeader(1));
            ClassicAssert.IsNull(policy.GetHeader(2));
            ClassicAssert.IsNull(policy.GetHeader(3));
            ClassicAssert.IsNull(policy.GetFooter(1));
            ClassicAssert.IsNull(policy.GetFooter(2));
            ClassicAssert.IsNull(policy.GetFooter(3));


            policy = header.GetHeaderFooterPolicy();
            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNull(policy.GetDefaultFooter());

            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(1));
            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(2));
            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(3));
            ClassicAssert.IsNull(policy.GetFooter(1));
            ClassicAssert.IsNull(policy.GetFooter(2));
            ClassicAssert.IsNull(policy.GetFooter(3));


            policy = footer.GetHeaderFooterPolicy();
            ClassicAssert.IsNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetDefaultFooter());

            ClassicAssert.IsNull(policy.GetHeader(1));
            ClassicAssert.IsNull(policy.GetHeader(2));
            ClassicAssert.IsNull(policy.GetHeader(3));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(1));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(2));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(3));


            policy = headerFooter.GetHeaderFooterPolicy();
            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetDefaultFooter());

            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(1));
            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(2));
            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(3));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(1));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(2));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(3));


            policy = oddEven.GetHeaderFooterPolicy();
            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetDefaultFooter());
            ClassicAssert.IsNotNull(policy.GetEvenPageHeader());
            ClassicAssert.IsNotNull(policy.GetEvenPageFooter());

            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(1));
            ClassicAssert.AreEqual(policy.GetEvenPageHeader(), policy.GetHeader(2));
            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(3));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(1));
            ClassicAssert.AreEqual(policy.GetEvenPageFooter(), policy.GetFooter(2));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(3));


            policy = diffFirst.GetHeaderFooterPolicy();
            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetDefaultFooter());
            ClassicAssert.IsNotNull(policy.GetFirstPageHeader());
            ClassicAssert.IsNotNull(policy.GetFirstPageFooter());
            ClassicAssert.IsNull(policy.GetEvenPageHeader());
            ClassicAssert.IsNull(policy.GetEvenPageFooter());

            ClassicAssert.AreEqual(policy.GetFirstPageHeader(), policy.GetHeader(1));
            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(2));
            ClassicAssert.AreEqual(policy.GetDefaultHeader(), policy.GetHeader(3));
            ClassicAssert.AreEqual(policy.GetFirstPageFooter(), policy.GetFooter(1));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(2));
            ClassicAssert.AreEqual(policy.GetDefaultFooter(), policy.GetFooter(3));
        }

        [Test]
        public void TestCreate()
        {
            XWPFDocument doc = new XWPFDocument();
            ClassicAssert.AreEqual(null, doc.GetHeaderFooterPolicy());
            ClassicAssert.AreEqual(0, doc.HeaderList.Count);
            ClassicAssert.AreEqual(0, doc.FooterList.Count);

            XWPFHeaderFooterPolicy policy = doc.CreateHeaderFooterPolicy();
            ClassicAssert.IsNotNull(doc.GetHeaderFooterPolicy());
            ClassicAssert.AreEqual(0, doc.HeaderList.Count);
            ClassicAssert.AreEqual(0, doc.FooterList.Count);

            // Create a header and a footer
            XWPFHeader header = policy.CreateHeader(XWPFHeaderFooterPolicy.DEFAULT);
            XWPFFooter footer = policy.CreateFooter(XWPFHeaderFooterPolicy.DEFAULT);
            header.CreateParagraph().CreateRun().SetText("Header Hello");
            footer.CreateParagraph().CreateRun().SetText("Footer Bye");


            // Save, re-load, and check
            doc = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            ClassicAssert.IsNotNull(doc.GetHeaderFooterPolicy());
            ClassicAssert.AreEqual(1, doc.HeaderList.Count);
            ClassicAssert.AreEqual(1, doc.FooterList.Count);

            ClassicAssert.AreEqual("Header Hello\n", doc.HeaderList[(0)].Text);
            ClassicAssert.AreEqual("Footer Bye\n", doc.FooterList[(0)].Text);
        }


        [Test]
        public void TestContents()
        {
            XWPFHeaderFooterPolicy policy;

            // Test a few simple bits off a simple header
            policy = diffFirst.GetHeaderFooterPolicy();

            ClassicAssert.AreEqual(
                "I am the header on the first page, and I" + '\u2019' + "m nice and simple\n",
                policy.GetFirstPageHeader().Text
            );
            ClassicAssert.AreEqual(
                    "First header column!\tMid header\tRight header!\n",
                    policy.GetDefaultHeader().Text
            );


            // And a few bits off a more complex header
            policy = oddEven.GetHeaderFooterPolicy();

            ClassicAssert.AreEqual(
                "[ODD Page Header text]\n\n",
                policy.GetDefaultHeader().Text
            );
            ClassicAssert.AreEqual(
                "[This is an Even Page, with a Header]\n\n",
                policy.GetEvenPageHeader().Text
            );
        }
    }

}
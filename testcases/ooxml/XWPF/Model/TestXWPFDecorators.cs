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
    using System.Linq;

    /**
     * Tests for the various XWPF decorators
     */
    [TestFixture]
    public class TestXWPFDecorators
    {
        private XWPFDocument simple;
        private XWPFDocument hyperlink;
        private XWPFDocument comments;
        private XWPFDocument footerhyperlink;
        private XWPFDocument footnotehyperlink;

        [SetUp]
        public void SetUp()
        {
            simple = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
            hyperlink = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
            comments = XWPFTestDataSamples.OpenSampleDocument("WordWithAttachments.docx");
            footerhyperlink = XWPFTestDataSamples.OpenSampleDocument("TestHyperlinkInFooterDocument.docx");
            footnotehyperlink = XWPFTestDataSamples.OpenSampleDocument("TestHyperlinkInFootnotes.docx");
        }

        [Test]
        public void TestHyperlink()
        {
            XWPFParagraph ps;
            XWPFParagraph ph;
            ClassicAssert.AreEqual(7, simple.Paragraphs.Count);
            ClassicAssert.AreEqual(5, hyperlink.Paragraphs.Count);

            // Simple text
            ps = simple.Paragraphs[(0)];
            ClassicAssert.AreEqual("I am a test document", ps.ParagraphText);
            ClassicAssert.AreEqual(1, ps.Runs.Count);

            ph = hyperlink.Paragraphs[(4)];
            ClassicAssert.AreEqual("We have a hyperlink here, and another.", ph.ParagraphText);
            ClassicAssert.AreEqual(3, ph.Runs.Count);


            // The proper way to do hyperlinks(!)
            ClassicAssert.IsFalse(ps.Runs[(0)] is XWPFHyperlinkRun);
            ClassicAssert.IsFalse(ph.Runs[(0)] is XWPFHyperlinkRun);
            ClassicAssert.IsTrue(ph.Runs[(1)] is XWPFHyperlinkRun);
            ClassicAssert.IsFalse(ph.Runs[(2)] is XWPFHyperlinkRun);

            XWPFHyperlinkRun link = (XWPFHyperlinkRun)ph.Runs[(1)];
            ClassicAssert.AreEqual("http://poi.apache.org/", link.GetHyperlink(hyperlink).URL);
        }

        [Test]
        public void TestHyperlinkInFooter()
        {
            ClassicAssert.AreEqual(1, footerhyperlink.Paragraphs.Count);

            // Simple text
            XWPFParagraph paragraph = footerhyperlink.Paragraphs[(0)];
            ClassicAssert.AreEqual("This is a test document.", paragraph.ParagraphText);
            ClassicAssert.AreEqual(2, paragraph.Runs.Count);
            
            ClassicAssert.AreEqual(3, footerhyperlink.FooterList.Count);

            ClassicAssert.AreEqual(1, footerhyperlink.GetHyperlinks().Length);

            XWPFHyperlinkRun run = (XWPFHyperlinkRun)((XWPFParagraph)footerhyperlink.FooterList[2].BodyElements[0]).Runs[1];
            ClassicAssert.AreEqual("http://poi.apache.org/", run.GetHyperlink(footerhyperlink).URL);

            ClassicAssert.AreEqual(1, footerhyperlink.FooterList[2].GetHyperlinks().Count);

            XWPFHyperlink link = footerhyperlink.GetHyperlinks().First();
            ClassicAssert.AreEqual("http://poi.apache.org/", link.URL);
        }

        [Test]
        public void TestHyperlinkInFootnotes()
        {
            ClassicAssert.AreEqual(1, footerhyperlink.Paragraphs.Count);

            // Simple text
            XWPFParagraph paragraph = footnotehyperlink.Paragraphs[(0)];
            ClassicAssert.AreEqual("This is a test document.[footnoteRef:1]", paragraph.ParagraphText);
            ClassicAssert.AreEqual(3, paragraph.Runs.Count);

            ClassicAssert.AreEqual(3, footnotehyperlink.GetFootnotes().Count);

            ClassicAssert.AreEqual(1, footnotehyperlink.GetHyperlinks().Length);

            XWPFHyperlinkRun run = (XWPFHyperlinkRun)footnotehyperlink.GetFootnotes()[2].Paragraphs[0].Runs[3];
            ClassicAssert.AreEqual("http://poi.apache.org/", run.GetHyperlink(footerhyperlink).URL);

            XWPFHyperlink link = footnotehyperlink.GetHyperlinks().First();
            ClassicAssert.AreEqual("http://poi.apache.org/", link.URL);
        }

        [Test]
        public void TestComments()
        {
            int numComments = 0;
            foreach (XWPFParagraph p in comments.Paragraphs)
            {
                XWPFCommentsDecorator d = new XWPFCommentsDecorator(p, null);
                if (d.GetCommentText().Length > 0)
                {
                    numComments++;
                    ClassicAssert.AreEqual("\tComment by", d.GetCommentText().Substring(0, 11));
                }
            }
            ClassicAssert.AreEqual(3, numComments);
        }
    }

}
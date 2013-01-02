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

namespace NPOI.XWPF.Model
{
    using System;



    using NUnit.Framework;

    using NPOI.XWPF;
    using NPOI.XWPF.UserModel;

    /**
     * Tests for the various XWPF decorators
     */
    [TestFixture]
    public class TestXWPFDecorators
    {
        private XWPFDocument simple;
        private XWPFDocument hyperlink;
        private XWPFDocument comments;
        [SetUp]
        public void SetUp()
        {
            simple = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
            hyperlink = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
            comments = XWPFTestDataSamples.OpenSampleDocument("WordWithAttachments.docx");
        }

        [Test]
        public void TestHyperlink()
        {
            XWPFParagraph ps;
            XWPFParagraph ph;
            Assert.AreEqual(7, simple.Paragraphs.Count);
            Assert.AreEqual(5, hyperlink.Paragraphs.Count);

            // Simple text
            ps = simple.Paragraphs[(0)];
            Assert.AreEqual("I am a test document", ps.GetParagraphText());
            Assert.AreEqual(1, ps.GetRuns().Count);

            ph = hyperlink.Paragraphs[(4)];
            Assert.AreEqual("We have a hyperlink here, and another.", ph.GetParagraphText());
            Assert.AreEqual(3, ph.GetRuns().Count);


            // The proper way to do hyperlinks(!)
            Assert.IsFalse(ps.GetRuns()[(0)] is XWPFHyperlinkRun);
            Assert.IsFalse(ph.GetRuns()[(0)] is XWPFHyperlinkRun);
            Assert.IsTrue(ph.GetRuns()[(1)] is XWPFHyperlinkRun);
            Assert.IsFalse(ph.GetRuns()[(2)] is XWPFHyperlinkRun);

            XWPFHyperlinkRun link = (XWPFHyperlinkRun)ph.GetRuns()[(1)];
            Assert.AreEqual("http://poi.apache.org/", link.GetHyperlink(hyperlink).URL);


            // Test the old style decorator
            // You probably don't want to still be using it...
            Assert.AreEqual(
                  "I am a test document",
                  (new XWPFHyperlinkDecorator(ps, null, false)).GetText()
            );
            Assert.AreEqual(
                  "I am a test document",
                  (new XWPFHyperlinkDecorator(ps, null, true)).GetText()
            );

            Assert.AreEqual(
                  "We have a hyperlink here, and another.hyperlink",
                  (new XWPFHyperlinkDecorator(ph, null, false)).GetText()
            );
            Assert.AreEqual(
                  "We have a hyperlink here, and another.hyperlink <http://poi.apache.org/>",
                  (new XWPFHyperlinkDecorator(ph, null, true)).GetText()
            );
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
                    Assert.AreEqual("\tComment by", d.GetCommentText().Substring(0, 11));
                }
            }
            Assert.AreEqual(3, numComments);
        }
    }

}
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

namespace NPOI.XWPF.UserModel
{
    using System;
    using NUnit.Framework;

    using NPOI.XWPF;
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.XWPF.Util;

    /**
     * Tests for XWPF Paragraphs
     */
    [TestFixture]
    public class TestXWPFParagraph
    {

        /**
         * Check that we Get the right paragraph from the header
         * @throws IOException 
         */
        [Test]
        public void TestHeaderParagraph()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");

            XWPFHeader hdr = xml.GetHeaderFooterPolicy().GetDefaultHeader();
            Assert.IsNotNull(hdr);

            IList<XWPFParagraph> ps = hdr.Paragraphs;
            Assert.AreEqual(1, ps.Count);
            XWPFParagraph p = ps[(0)];

            Assert.AreEqual(5, p.GetCTP().GetRList().Count);
            Assert.AreEqual("First header column!\tMid header\tRight header!", p.Text);
        }

        /**
         * Check that we Get the right paragraphs from the document
         * @throws IOException 
         */
        [Test]
        public void TestDocumentParagraph()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");
            IList<XWPFParagraph> ps = xml.Paragraphs;
            Assert.AreEqual(10, ps.Count);

            Assert.IsFalse(ps[(0)].IsEmpty);
            Assert.AreEqual(
                    "This is a sample word document. It has two pages. It has a three column heading, but no footer.",
                    ps[(0)].Text);

            Assert.IsTrue(ps[1].IsEmpty);
            Assert.AreEqual("", ps[1].Text);

            Assert.IsFalse(ps[2].IsEmpty);
            Assert.AreEqual("HEADING TEXT", ps[2].Text);

            Assert.IsTrue(ps[3].IsEmpty);
            Assert.AreEqual("", ps[3].Text);

            Assert.IsFalse(ps[4].IsEmpty);
            Assert.AreEqual("More on page one", ps[4].Text);
        }

        [Test]
        public void TestSetBorderTop()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            Assert.AreEqual(ST_Border.none, EnumConverter.ValueOf<ST_Border, Borders>(p.BorderTop));

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            //bordi
            CT_PBdr bdr = ppr.AddNewPBdr();
            CT_Border borderTop = bdr.AddNewTop();
            borderTop.val = (ST_Border.@double);
            bdr.top = (borderTop);

            Assert.AreEqual(Borders.DOUBLE, p.BorderTop);
            p.BorderTop = (Borders.SINGLE);
            Assert.AreEqual(ST_Border.single, borderTop.val);
        }

        [Test]
        public void TestSetAlignment()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            Assert.AreEqual(ParagraphAlignment.LEFT, p.Alignment);

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            CT_Jc align = ppr.AddNewJc();
            align.val = (ST_Jc.center);
            Assert.AreEqual(ParagraphAlignment.CENTER, p.Alignment);

            p.Alignment = (ParagraphAlignment.BOTH);
            Assert.AreEqual((int)ST_Jc.both, (int)ppr.jc.val);
        }


        [Test]
        public void TestSetGetSpacing()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            Assert.AreEqual(-1, p.SpacingAfter);

            CT_Spacing spacing = ppr.AddNewSpacing();
            spacing.after = 10;
            Assert.AreEqual(10, p.SpacingAfter);

            p.SpacingAfter = 100;
            Assert.AreEqual(100, (int)spacing.after);
        }

        [Test]
        public void TestSetGetSpacingLineRule()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            Assert.AreEqual(LineSpacingRule.AUTO, p.SpacingLineRule);

            CT_Spacing spacing = ppr.AddNewSpacing();
            spacing.lineRule = (ST_LineSpacingRule.atLeast);
            Assert.AreEqual(LineSpacingRule.ATLEAST, p.SpacingLineRule);

            p.SpacingAfter = 100;
            Assert.AreEqual(100, (int)spacing.after);
        }

        [Test]
        public void TestSetGetIndentation()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            Assert.AreEqual(-1, p.IndentationLeft);

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            Assert.AreEqual(-1, p.IndentationLeft);

            CT_Ind ind = ppr.AddNewInd();
            ind.left = "10";
            Assert.AreEqual(10, p.IndentationLeft);

            p.IndentationLeft = 100;
            Assert.AreEqual(100, int.Parse(ind.left));
        }

        [Test]
        public void TestSetGetVerticalAlignment()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            CT_TextAlignment txtAlign = ppr.AddNewTextAlignment();
            txtAlign.val = (ST_TextAlignment.center);
            Assert.AreEqual(TextAlignment.CENTER, p.VerticalAlignment);

            p.VerticalAlignment= (TextAlignment.BOTTOM);
            Assert.AreEqual(ST_TextAlignment.bottom, ppr.textAlignment.val);
        }

        [Test]
        public void TestSetGetWordWrap()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            CT_OnOff wordWrap = ppr.AddNewWordWrap();
            wordWrap.val = false;
            Assert.AreEqual(false, p.IsWordWrap);

            p.IsWordWrap = true;
            Assert.AreEqual(true, ppr.wordWrap.val);
        }


        [Test]
        public void TestSetGetPageBreak()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            CT_OnOff pageBreak = ppr.AddNewPageBreakBefore();
            pageBreak.val = false;
            Assert.AreEqual(false, p.IsPageBreak);

            p.IsPageBreak = (true);
            Assert.AreEqual(true, ppr.pageBreakBefore.val);
        }

        [Test]
        public void TestBookmarks()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("bookmarks.docx");
            XWPFParagraph paragraph = doc.Paragraphs[0];
            Assert.AreEqual("Sample Word Document", paragraph.Text);
            Assert.AreEqual(1, paragraph.GetCTP().SizeOfBookmarkStartArray());
            Assert.AreEqual(0, paragraph.GetCTP().SizeOfBookmarkEndArray());
            CT_Bookmark ctBookmark = paragraph.GetCTP().GetBookmarkStartArray(0);
            Assert.AreEqual("poi", ctBookmark.name);
            foreach (CT_Bookmark bookmark in paragraph.GetCTP().GetBookmarkStartList())
            {
                Assert.AreEqual("poi", bookmark.name);
            }
        }

        [Test]
        public void TestGetSetNumID()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            p.SetNumID ("10");
            Assert.AreEqual("10", p.GetNumID());
        }

        [Test]
        public void TestAddingRuns()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");

            XWPFParagraph p = doc.Paragraphs[0];
            Assert.AreEqual(2, p.Runs.Count);

            XWPFRun r = p.CreateRun();
            Assert.AreEqual(3, p.Runs.Count);
            Assert.AreEqual(2, p.Runs.IndexOf(r));

            XWPFRun r2 = p.InsertNewRun(1);
            Assert.AreEqual(4, p.Runs.Count);
            Assert.AreEqual(1, p.Runs.IndexOf(r2));
            Assert.AreEqual(3, p.Runs.IndexOf(r));
        }

        [Test]
        public void TestPictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("VariousPictures.docx");
            Assert.AreEqual(7, doc.Paragraphs.Count);

            XWPFParagraph p;
            XWPFRun r;

            // Text paragraphs
            Assert.AreEqual("Sheet with various pictures", doc.Paragraphs[0].Text);
            Assert.AreEqual("(jpeg, png, wmf, emf and pict) ", doc.Paragraphs[1].Text);

            // Spacer ones
            Assert.AreEqual("", doc.Paragraphs[2].Text);
            Assert.AreEqual("", doc.Paragraphs[3].Text);
            Assert.AreEqual("", doc.Paragraphs[4].Text);

            // Image one
            p = doc.Paragraphs[5];
            Assert.AreEqual(6, p.Runs.Count);

            r = p.Runs[0];
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.GetEmbeddedPictures().Count);
            Assert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            Assert.AreEqual("image1.wmf", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[1];
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.GetEmbeddedPictures().Count);
            Assert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            Assert.AreEqual("image2.png", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[2];
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.GetEmbeddedPictures().Count);
            Assert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            Assert.AreEqual("image3.emf", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[3];
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.GetEmbeddedPictures().Count);
            Assert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            Assert.AreEqual("image4.emf", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[4];
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.GetEmbeddedPictures().Count);
            Assert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            Assert.AreEqual("image5.jpeg", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[5];
            //Is there a bug about XmlSerializer? it can not Deserialize the tag which inner text is only one whitespace
            //e.g. <w:t> </w:t> to CT_Text;
            //TODO 
            Assert.AreEqual(" ", r.ToString());
            Assert.AreEqual(0, r.GetEmbeddedPictures().Count);

            // Final spacer
            Assert.AreEqual("", doc.Paragraphs[(6)].Text);


            // Look in detail at one
            r = p.Runs[4];
            XWPFPicture pict = r.GetEmbeddedPictures()[0];
            //CT_Picture picture = pict.GetCTPicture();
            NPOI.OpenXmlFormats.Dml.Picture.CT_Picture picture = pict.GetCTPicture();
            //Assert.Fail("picture.blipFill.blip.embed is missing from wordprocessing CT_Picture.");
            Assert.AreEqual("rId8", picture.blipFill.blip.embed);

            // Ensure that the ooxml compiler Finds everything we need
            r.GetCTR().GetDrawingArray(0);
            r.GetCTR().GetDrawingArray(0).GetInlineArray(0);
            NPOI.OpenXmlFormats.Dml.CT_GraphicalObject go = r.GetCTR().GetDrawingArray(0).GetInlineArray(0).graphic;
            NPOI.OpenXmlFormats.Dml.CT_GraphicalObjectData god =  r.GetCTR().GetDrawingArray(0).GetInlineArray(0).graphic.graphicData;
            //PicDocument pd = new PicDocumentImpl(null);
        }
    }

}
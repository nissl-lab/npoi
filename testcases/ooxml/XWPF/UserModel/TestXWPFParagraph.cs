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





    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.XWPF;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
    using org.Openxmlformats.schemas.Drawingml.x2006.picture;
    using org.Openxmlformats.schemas.Drawingml.x2006.picture;
    using org.Openxmlformats.schemas.Drawingml.x2006.picture.impl;

    /**
     * Tests for XWPF Paragraphs
     */
    [TestClass]
    public class TestXWPFParagraph
    {

        /**
         * Check that we Get the right paragraph from the header
         * @throws IOException 
         */
        [TestMethod]
        public void TestHeaderParagraph()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");

            XWPFHeader hdr = xml.HeaderFooterPolicy.DefaultHeader;
            Assert.IsNotNull(hdr);

            List<XWPFParagraph> ps = hdr.Paragraphs;
            Assert.AreEqual(1, ps.Size());
            XWPFParagraph p = ps.Get(0);

            Assert.AreEqual(5, p.CTP.RList.Size());
            Assert.AreEqual("First header column!\tMid header\tRight header!", p
                    .Text);
        }

        /**
         * Check that we Get the right paragraphs from the document
         * @throws IOException 
         */
        [TestMethod]
        public void TestDocumentParagraph()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");
            List<XWPFParagraph> ps = xml.Paragraphs;
            Assert.AreEqual(10, ps.Size());

            Assert.IsFalse(ps.Get(0).IsEmpty());
            Assert.AreEqual(
                    "This is a sample word document. It has two pages. It has a three column heading, but no footer.",
                    ps.Get(0).Text);

            Assert.IsTrue(ps.Get(1).IsEmpty());
            Assert.AreEqual("", ps.Get(1).Text);

            Assert.IsFalse(ps.Get(2).IsEmpty());
            Assert.AreEqual("HEADING TEXT", ps.Get(2).Text);

            Assert.IsTrue(ps.Get(3).IsEmpty());
            Assert.AreEqual("", ps.Get(3).Text);

            Assert.IsFalse(ps.Get(4).IsEmpty());
            Assert.AreEqual("More on page one", ps.Get(4).Text);
        }

        [TestMethod]
        public void TestSetGetBorderTop()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            Assert.AreEqual(STBorder.NONE.IntValue(), p.BorderTop.Value);

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            //bordi
            CTPBdr bdr = ppr.AddNewPBdr();
            CTBorder borderTop = bdr.AddNewTop();
            borderTop.Val = (STBorder.DOUBLE);
            bdr.Top = (borderTop);

            Assert.AreEqual(Borders.DOUBLE, p.BorderTop);
            p.BorderTop = (Borders.SINGLE);
            Assert.AreEqual(STBorder.SINGLE, borderTop.Val);
        }

        [TestMethod]
        public void TestSetGetAlignment()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            Assert.AreEqual(STJc.LEFT.IntValue(), p.Alignment.Value);

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            CTJc align = ppr.AddNewJc();
            align.Val = (STJc.CENTER);
            Assert.AreEqual(ParagraphAlignment.CENTER, p.Alignment);

            p.Alignment = (ParagraphAlignment.BOTH);
            Assert.AreEqual(STJc.BOTH, ppr.Jc.Val);
        }


        [TestMethod]
        public void TestSetGetSpacing()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            Assert.AreEqual(-1, p.SpacingAfter);

            CTSpacing spacing = ppr.AddNewSpacing();
            spacing.After = (new Bigint("10"));
            Assert.AreEqual(10, p.SpacingAfter);

            p.SpacingAfter = (100);
            Assert.AreEqual(100, spacing.After.IntValue());
        }

        [TestMethod]
        public void TestSetGetSpacingLineRule()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            Assert.AreEqual(STLineSpacingRule.INT_AUTO, p.SpacingLineRule.Value);

            CTSpacing spacing = ppr.AddNewSpacing();
            spacing.LineRule = (STLineSpacingRule.AT_LEAST);
            Assert.AreEqual(LineSpacingRule.AT_LEAST, p.SpacingLineRule);

            p.SpacingAfter = (100);
            Assert.AreEqual(100, spacing.After.IntValue());
        }

        [TestMethod]
        public void TestSetGetIndentation()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            Assert.AreEqual(-1, p.IndentationLeft);

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            Assert.AreEqual(-1, p.IndentationLeft);

            CTInd ind = ppr.AddNewInd();
            ind.Left = (new Bigint("10"));
            Assert.AreEqual(10, p.IndentationLeft);

            p.IndentationLeft = (100);
            Assert.AreEqual(100, ind.Left.IntValue());
        }

        [TestMethod]
        public void TestSetGetVerticalAlignment()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            CTTextAlignment txtAlign = ppr.AddNewTextAlignment();
            txtAlign.Val = (STTextAlignment.CENTER);
            Assert.AreEqual(TextAlignment.CENTER, p.VerticalAlignment);

            p.VerticalAlignment = (TextAlignment.BOTTOM);
            Assert.AreEqual(STTextAlignment.BOTTOM, ppr.TextAlignment.Val);
        }

        [TestMethod]
        public void TestSetGetWordWrap()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            CTOnOff wordWrap = ppr.AddNewWordWrap();
            wordWrap.Val = (STOnOff.FALSE);
            Assert.AreEqual(false, p.IsWordWrap());

            p.WordWrap = (true);
            Assert.AreEqual(STOnOff.TRUE, ppr.WordWrap.Val);
        }


        [TestMethod]
        public void TestSetGetPageBreak()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CTP ctp = p.CTP;
            CTPPr ppr = ctp.PPr == null ? ctp.AddNewPPr() : ctp.PPr;

            CTOnOff pageBreak = ppr.AddNewPageBreakBefore();
            pageBreak.Val = (STOnOff.FALSE);
            Assert.AreEqual(false, p.IsPageBreak());

            p.PageBreak = (true);
            Assert.AreEqual(STOnOff.TRUE, ppr.PageBreakBefore.Val);
        }

        [TestMethod]
        public void TestBookmarks()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("bookmarks.docx");
            XWPFParagraph paragraph = doc.Paragraphs.Get(0);
            Assert.AreEqual("Sample Word Document", paragraph.Text);
            Assert.AreEqual(1, paragraph.CTP.SizeOfBookmarkStartArray());
            Assert.AreEqual(0, paragraph.CTP.SizeOfBookmarkEndArray());
            CTBookmark ctBookmark = paragraph.CTP.GetBookmarkStartArray(0);
            Assert.AreEqual("poi", ctBookmark.Name);
            foreach (CTBookmark bookmark in paragraph.CTP.BookmarkStartList)
            {
                Assert.AreEqual("poi", bookmark.Name);
            }
        }

        [TestMethod]
        public void TestGetSetNumID()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            p.NumID = (new Bigint("10"));
            Assert.AreEqual("10", p.NumID.ToString());
        }

        [TestMethod]
        public void TestAddingRuns()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");

            XWPFParagraph p = doc.Paragraphs.Get(0);
            Assert.AreEqual(2, p.Runs.Size());

            XWPFRun r = p.CreateRun();
            Assert.AreEqual(3, p.Runs.Size());
            Assert.AreEqual(2, p.Runs.IndexOf(r));

            XWPFRun r2 = p.InsertNewRun(1);
            Assert.AreEqual(4, p.Runs.Size());
            Assert.AreEqual(1, p.Runs.IndexOf(r2));
            Assert.AreEqual(3, p.Runs.IndexOf(r));
        }

        [TestMethod]
        public void TestPictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("VariousPictures.docx");
            Assert.AreEqual(7, doc.Paragraphs.Size());

            XWPFParagraph p;
            XWPFRun r;

            // Text paragraphs
            Assert.AreEqual("Sheet with various pictures", doc.Paragraphs.Get(0).Text);
            Assert.AreEqual("(jpeg, png, wmf, emf and pict) ", doc.Paragraphs.Get(1).Text);

            // Spacer ones
            Assert.AreEqual("", doc.Paragraphs.Get(2).Text);
            Assert.AreEqual("", doc.Paragraphs.Get(3).Text);
            Assert.AreEqual("", doc.Paragraphs.Get(4).Text);

            // Image one
            p = doc.Paragraphs.Get(5);
            Assert.AreEqual(6, p.Runs.Size());

            r = p.Runs.Get(0);
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.EmbeddedPictures.Size());
            Assert.IsNotNull(r.EmbeddedPictures.Get(0).PictureData);
            Assert.AreEqual("image1.wmf", r.EmbeddedPictures.Get(0).PictureData.FileName);

            r = p.Runs.Get(1);
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.EmbeddedPictures.Size());
            Assert.IsNotNull(r.EmbeddedPictures.Get(0).PictureData);
            Assert.AreEqual("image2.png", r.EmbeddedPictures.Get(0).PictureData.FileName);

            r = p.Runs.Get(2);
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.EmbeddedPictures.Size());
            Assert.IsNotNull(r.EmbeddedPictures.Get(0).PictureData);
            Assert.AreEqual("image3.emf", r.EmbeddedPictures.Get(0).PictureData.FileName);

            r = p.Runs.Get(3);
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.EmbeddedPictures.Size());
            Assert.IsNotNull(r.EmbeddedPictures.Get(0).PictureData);
            Assert.AreEqual("image4.emf", r.EmbeddedPictures.Get(0).PictureData.FileName);

            r = p.Runs.Get(4);
            Assert.AreEqual("", r.ToString());
            Assert.AreEqual(1, r.EmbeddedPictures.Size());
            Assert.IsNotNull(r.EmbeddedPictures.Get(0).PictureData);
            Assert.AreEqual("image5.jpeg", r.EmbeddedPictures.Get(0).PictureData.FileName);

            r = p.Runs.Get(5);
            Assert.AreEqual(" ", r.ToString());
            Assert.AreEqual(0, r.EmbeddedPictures.Size());

            // Final spacer
            Assert.AreEqual("", doc.Paragraphs.Get(6).Text);


            // Look in detail at one
            r = p.Runs.Get(4);
            XWPFPicture pict = r.EmbeddedPictures.Get(0);
            CTPicture picture = pict.CTPicture;
            Assert.AreEqual("rId8", picture.BlipFill.Blip.Embed);

            // Ensure that the ooxml compiler Finds everything we need
            r.CTR.GetDrawingArray(0);
            r.CTR.GetDrawingArray(0).GetInlineArray(0);
            r.CTR.GetDrawingArray(0).GetInlineArray(0).Graphic;
            r.CTR.GetDrawingArray(0).GetInlineArray(0).Graphic.GraphicData;
            PicDocument pd = new PicDocumentImpl(null);
        }
    }

}
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

namespace TestCases.XWPF.UserModel
{
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using System.Collections.Generic;
    using System.Text;

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
            ClassicAssert.IsNotNull(hdr);

            IList<XWPFParagraph> ps = hdr.Paragraphs;
            ClassicAssert.AreEqual(1, ps.Count);
            XWPFParagraph p = ps[(0)];

            ClassicAssert.AreEqual(5, p.GetCTP().GetRList().Count);
            ClassicAssert.AreEqual("First header column!\tMid header\tRight header!", p.Text);
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
            ClassicAssert.AreEqual(10, ps.Count);

            ClassicAssert.IsFalse(ps[(0)].IsEmpty);
            ClassicAssert.AreEqual(
                    "This is a sample word document. It has two pages. It has a three column heading, but no footer.",
                    ps[(0)].Text);

            ClassicAssert.IsTrue(ps[1].IsEmpty);
            ClassicAssert.AreEqual("", ps[1].Text);

            ClassicAssert.IsFalse(ps[2].IsEmpty);
            ClassicAssert.AreEqual("HEADING TEXT", ps[2].Text);

            ClassicAssert.IsTrue(ps[3].IsEmpty);
            ClassicAssert.AreEqual("", ps[3].Text);

            ClassicAssert.IsFalse(ps[4].IsEmpty);
            ClassicAssert.AreEqual("More on page one", ps[4].Text);
        }

        [Test]
        public void TestSetBorderTop()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            ClassicAssert.AreEqual(ST_Border.none, EnumConverter.ValueOf<ST_Border, Borders>(p.BorderTop));

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            //bordi
            CT_PBdr bdr = ppr.AddNewPBdr();
            CT_Border borderTop = bdr.AddNewTop();
            borderTop.val = (ST_Border.@double);
            bdr.top = (borderTop);

            ClassicAssert.AreEqual(Borders.Double, p.BorderTop);
            p.BorderTop = (Borders.Single);
            ClassicAssert.AreEqual(ST_Border.single, borderTop.val);
        }

        [Test]
        public void TestSetAlignment()
        {
            //new clean instance of paragraph
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            ClassicAssert.AreEqual(ParagraphAlignment.LEFT, p.Alignment);

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            CT_Jc align = ppr.AddNewJc();
            align.val = (ST_Jc.center);
            ClassicAssert.AreEqual(ParagraphAlignment.CENTER, p.Alignment);

            p.Alignment = (ParagraphAlignment.BOTH);
            ClassicAssert.AreEqual((int)ST_Jc.both, (int)ppr.jc.val);
        }

        [Test]
        public void TestSetGetSpacing()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            ClassicAssert.AreEqual(-1, p.SpacingBefore);
            ClassicAssert.AreEqual(-1, p.SpacingAfter);
            ClassicAssert.AreEqual(-1, p.SpacingBetween, 0.1);
            ClassicAssert.AreEqual(LineSpacingRule.AUTO, p.SpacingLineRule);

            CT_Spacing spacing = ppr.AddNewSpacing();
            spacing.after = 10;
            ClassicAssert.AreEqual(10, p.SpacingAfter);
            spacing.before = 10;
            ClassicAssert.AreEqual(10, p.SpacingBefore);

            p.SpacingAfter = 100;
            ClassicAssert.AreEqual(100, (int)spacing.after);
            p.SpacingBefore = 100;
            ClassicAssert.AreEqual(100, spacing.before);

            p.SetSpacingBetween(.25, LineSpacingRule.EXACT);
            ClassicAssert.AreEqual(.25, p.SpacingBetween, 0.01);
            ClassicAssert.AreEqual(LineSpacingRule.EXACT, p.SpacingLineRule);
            p.SetSpacingBetween(1.25, LineSpacingRule.AUTO);
            ClassicAssert.AreEqual(1.25, p.SpacingBetween, 0.01);
            ClassicAssert.AreEqual(LineSpacingRule.AUTO, p.SpacingLineRule);
            p.SetSpacingBetween(.5, LineSpacingRule.ATLEAST);
            ClassicAssert.AreEqual(.5, p.SpacingBetween, 0.01);
            ClassicAssert.AreEqual(LineSpacingRule.ATLEAST, p.SpacingLineRule);
            p.SetSpacingBetween(1.15);
            ClassicAssert.AreEqual(1.15, p.SpacingBetween, 0.01);
            ClassicAssert.AreEqual(LineSpacingRule.AUTO, p.SpacingLineRule);

            doc.Close();
        }

        [Test]
        public void TestSetGetSpacingLineRule()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            ClassicAssert.AreEqual(LineSpacingRule.AUTO, p.SpacingLineRule);

            CT_Spacing spacing = ppr.AddNewSpacing();
            spacing.lineRule = (ST_LineSpacingRule.atLeast);
            ClassicAssert.AreEqual(LineSpacingRule.ATLEAST, p.SpacingLineRule);

            p.SpacingAfter = 100;
            ClassicAssert.AreEqual(100, (int)spacing.after);
        }

        [Test]
        public void TestSetGetIndentation()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            ClassicAssert.AreEqual(-1, p.IndentationLeft);

            CT_P ctp = p.GetCTP();
            CT_PPr ppr = ctp.pPr == null ? ctp.AddNewPPr() : ctp.pPr;

            ClassicAssert.AreEqual(-1, p.IndentationLeft);

            CT_Ind ind = ppr.AddNewInd();
            ind.left = "10";
            ClassicAssert.AreEqual(10, p.IndentationLeft);

            p.IndentationLeft = 100;
            ClassicAssert.AreEqual(100, int.Parse(ind.left));
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
            ClassicAssert.AreEqual(TextAlignment.CENTER, p.VerticalAlignment);

            p.VerticalAlignment = (TextAlignment.BOTTOM);
            ClassicAssert.AreEqual(ST_TextAlignment.bottom, ppr.textAlignment.val);
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
            ClassicAssert.AreEqual(false, p.IsWordWrapped);

            p.IsWordWrapped = true;
            ClassicAssert.AreEqual(true, ppr.wordWrap.val);
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
            ClassicAssert.AreEqual(false, p.IsPageBreak);

            p.IsPageBreak = (true);
            ClassicAssert.AreEqual(true, ppr.pageBreakBefore.val);
        }

        [Test]
        public void TestBookmarks()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("bookmarks.docx");
            XWPFParagraph paragraph = doc.Paragraphs[0];
            ClassicAssert.AreEqual("Sample Word Document", paragraph.Text);
            ClassicAssert.AreEqual(1, paragraph.GetCTP().SizeOfBookmarkStartArray());
            ClassicAssert.AreEqual(0, paragraph.GetCTP().SizeOfBookmarkEndArray());
            CT_Bookmark ctBookmark = paragraph.GetCTP().GetBookmarkStartArray(0);
            ClassicAssert.AreEqual("poi", ctBookmark.name);
            foreach (CT_Bookmark bookmark in paragraph.GetCTP().GetBookmarkStartList())
            {
                ClassicAssert.AreEqual("poi", bookmark.name);
            }
        }

        [Test]
        public void TestGetSetNumID()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            p.SetNumID("10");
            ClassicAssert.AreEqual("10", p.GetNumID());
        }

        [Test]
        public void TestGetSetILvl()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            p.SetNumILvl("1");
            ClassicAssert.AreEqual("1", p.GetNumIlvl());

        }

        [Test]
        public void TestAddingRuns()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");

            XWPFParagraph p = doc.Paragraphs[0];
            ClassicAssert.AreEqual(2, p.Runs.Count);

            XWPFRun r = p.CreateRun();
            ClassicAssert.AreEqual(3, p.Runs.Count);
            ClassicAssert.AreEqual(2, p.Runs.IndexOf(r));

            XWPFRun r2 = p.InsertNewRun(1);
            ClassicAssert.AreEqual(4, p.Runs.Count);
            ClassicAssert.AreEqual(1, p.Runs.IndexOf(r2));
            ClassicAssert.AreEqual(3, p.Runs.IndexOf(r));
        }

        [Test]
        public void TestPictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("VariousPictures.docx");
            ClassicAssert.AreEqual(7, doc.Paragraphs.Count);

            XWPFParagraph p;
            XWPFRun r;

            // Text paragraphs
            ClassicAssert.AreEqual("Sheet with various pictures", doc.Paragraphs[0].Text);
            ClassicAssert.AreEqual("(jpeg, png, wmf, emf and pict) ", doc.Paragraphs[1].Text);

            // Spacer ones
            ClassicAssert.AreEqual("", doc.Paragraphs[2].Text);
            ClassicAssert.AreEqual("", doc.Paragraphs[3].Text);
            ClassicAssert.AreEqual("", doc.Paragraphs[4].Text);

            // Image one
            p = doc.Paragraphs[5];
            ClassicAssert.AreEqual(6, p.Runs.Count);

            r = p.Runs[0];
            ClassicAssert.AreEqual("", r.ToString());
            ClassicAssert.AreEqual(1, r.GetEmbeddedPictures().Count);
            ClassicAssert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            ClassicAssert.AreEqual("image1.wmf", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[1];
            ClassicAssert.AreEqual("", r.ToString());
            ClassicAssert.AreEqual(1, r.GetEmbeddedPictures().Count);
            ClassicAssert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            ClassicAssert.AreEqual("image2.png", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[2];
            ClassicAssert.AreEqual("", r.ToString());
            ClassicAssert.AreEqual(1, r.GetEmbeddedPictures().Count);
            ClassicAssert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            ClassicAssert.AreEqual("image3.emf", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[3];
            ClassicAssert.AreEqual("", r.ToString());
            ClassicAssert.AreEqual(1, r.GetEmbeddedPictures().Count);
            ClassicAssert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            ClassicAssert.AreEqual("image4.emf", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[4];
            ClassicAssert.AreEqual("", r.ToString());
            ClassicAssert.AreEqual(1, r.GetEmbeddedPictures().Count);
            ClassicAssert.IsNotNull(r.GetEmbeddedPictures()[0].GetPictureData());
            ClassicAssert.AreEqual("image5.jpeg", r.GetEmbeddedPictures()[0].GetPictureData().FileName);

            r = p.Runs[5];
            //Is there a bug about XmlSerializer? it can not Deserialize the tag which inner text is only one whitespace
            //e.g. <w:t> </w:t> to CT_Text;
            //TODO 
            ClassicAssert.AreEqual(" ", r.ToString());
            ClassicAssert.AreEqual(0, r.GetEmbeddedPictures().Count);

            // Final spacer
            ClassicAssert.AreEqual("", doc.Paragraphs[(6)].Text);


            // Look in detail at one
            r = p.Runs[4];
            XWPFPicture pict = r.GetEmbeddedPictures()[0];
            //CT_Picture picture = pict.GetCTPicture();
            NPOI.OpenXmlFormats.Dml.Picture.CT_Picture picture = pict.GetCTPicture();
            //Assert.Fail("picture.blipFill.blip.embed is missing from wordprocessing CT_Picture.");
            ClassicAssert.AreEqual("rId8", picture.blipFill.blip.embed);

            // Ensure that the ooxml compiler Finds everything we need
            r.GetCTR().GetDrawingArray(0);
            r.GetCTR().GetDrawingArray(0).GetInlineArray(0);
            NPOI.OpenXmlFormats.Dml.CT_GraphicalObject go = r.GetCTR().GetDrawingArray(0).GetInlineArray(0).graphic;
            NPOI.OpenXmlFormats.Dml.CT_GraphicalObjectData god = r.GetCTR().GetDrawingArray(0).GetInlineArray(0).graphic.graphicData;
            //PicDocument pd = new PicDocumentImpl(null);
            //assertTrue(pd.isNil());
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestTika792()
        {
            //This test forces the loading of CTMoveBookmark and
            //CTMoveBookmarkImpl into ooxml-lite.
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Tika-792.docx");
            XWPFParagraph paragraph = doc.Paragraphs[(0)];
            ClassicAssert.AreEqual("", paragraph.Text);
            paragraph = doc.Paragraphs[1];
            ClassicAssert.AreEqual("b", paragraph.Text);
        }

        [Test]
        public void TestSettersGetters()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            //ClassicAssert.IsTrue(p.IsEmpty);
            ClassicAssert.IsFalse(p.RemoveRun(0));

            p.BorderTop = (/*setter*/Borders.BabyPacifier);
            p.BorderBetween = (/*setter*/Borders.BabyPacifier);
            p.BorderBottom = (/*setter*/Borders.BabyPacifier);

            ClassicAssert.IsNotNull(p.IRuns);
            ClassicAssert.AreEqual(0, p.IRuns.Count);
            ClassicAssert.IsFalse(p.IsEmpty);
            ClassicAssert.IsNull(p.StyleID);
            ClassicAssert.IsNull(p.Style);

            ClassicAssert.IsNull(p.GetNumID());
            p.SetNumID("12");
            ClassicAssert.AreEqual("12", p.GetNumID());
            p.SetNumID("13");
            ClassicAssert.AreEqual("13", p.GetNumID());

            ClassicAssert.IsNull(p.GetNumFmt());

            //ClassicAssert.IsNull(p.GetNumIlvl());

            ClassicAssert.AreEqual("", p.ParagraphText);
            ClassicAssert.AreEqual("", p.PictureText);
            ClassicAssert.AreEqual("", p.FootnoteText);

            p.BorderBetween = (/*setter*/Borders.None);
            ClassicAssert.AreEqual(Borders.None, p.BorderBetween);
            p.BorderBetween = (/*setter*/Borders.BasicBlackDashes);
            ClassicAssert.AreEqual(Borders.BasicBlackDashes, p.BorderBetween);

            p.BorderBottom = (/*setter*/Borders.None);
            ClassicAssert.AreEqual(Borders.None, p.BorderBottom);
            p.BorderBottom = (/*setter*/Borders.BabyPacifier);
            ClassicAssert.AreEqual(Borders.BabyPacifier, p.BorderBottom);

            p.BorderLeft = (/*setter*/Borders.None);
            ClassicAssert.AreEqual(Borders.None, p.BorderLeft);
            p.BorderLeft = (/*setter*/Borders.BasicWhiteSquares);
            ClassicAssert.AreEqual(Borders.BasicWhiteSquares, p.BorderLeft);

            p.BorderRight = (/*setter*/Borders.None);
            ClassicAssert.AreEqual(Borders.None, p.BorderRight);
            p.BorderRight = (/*setter*/Borders.BasicWhiteDashes);
            ClassicAssert.AreEqual(Borders.BasicWhiteDashes, p.BorderRight);

            p.BorderBottom = (/*setter*/Borders.None);
            ClassicAssert.AreEqual(Borders.None, p.BorderBottom);
            p.BorderBottom = (/*setter*/Borders.BasicWhiteDots);
            ClassicAssert.AreEqual(Borders.BasicWhiteDots, p.BorderBottom);

            ClassicAssert.IsFalse(p.IsPageBreak);
            p.IsPageBreak = (/*setter*/true);
            ClassicAssert.IsTrue(p.IsPageBreak);
            p.IsPageBreak = (/*setter*/false);
            ClassicAssert.IsFalse(p.IsPageBreak);

            ClassicAssert.AreEqual(-1, p.SpacingAfter);
            p.SpacingAfter = (/*setter*/12);
            ClassicAssert.AreEqual(12, p.SpacingAfter);

            ClassicAssert.AreEqual(-1, p.SpacingAfterLines);
            p.SpacingAfterLines = (/*setter*/14);
            ClassicAssert.AreEqual(14, p.SpacingAfterLines);

            ClassicAssert.AreEqual(-1, p.SpacingBefore);
            p.SpacingBefore = (/*setter*/16);
            ClassicAssert.AreEqual(16, p.SpacingBefore);

            ClassicAssert.AreEqual(-1, p.SpacingBeforeLines);
            p.SpacingBeforeLines = (/*setter*/18);
            ClassicAssert.AreEqual(18, p.SpacingBeforeLines);

            ClassicAssert.AreEqual(LineSpacingRule.AUTO, p.SpacingLineRule);
            p.SpacingLineRule = (/*setter*/LineSpacingRule.EXACT);
            ClassicAssert.AreEqual(LineSpacingRule.EXACT, p.SpacingLineRule);

            ClassicAssert.AreEqual(-1, p.IndentationLeft);
            p.IndentationLeft = (/*setter*/21);
            ClassicAssert.AreEqual(21, p.IndentationLeft);

            ClassicAssert.AreEqual(-1, p.IndentationRight);
            p.IndentationRight = (/*setter*/25);
            ClassicAssert.AreEqual(25, p.IndentationRight);

            ClassicAssert.AreEqual(-1, p.IndentationHanging);
            p.IndentationHanging = (/*setter*/25);
            ClassicAssert.AreEqual(25, p.IndentationHanging);

            ClassicAssert.AreEqual(-1, p.IndentationFirstLine);
            p.IndentationFirstLine = (/*setter*/25);
            ClassicAssert.AreEqual(25, p.IndentationFirstLine);

            ClassicAssert.IsFalse(p.IsWordWrapped);
            p.IsWordWrapped = (/*setter*/true);
            ClassicAssert.IsTrue(p.IsWordWrapped);
            p.IsWordWrapped = (/*setter*/false);
            ClassicAssert.IsFalse(p.IsWordWrapped);

            ClassicAssert.IsNull(p.Style);
            p.Style = (/*setter*/"teststyle");
            ClassicAssert.AreEqual("teststyle", p.Style);

            p.AddRun(new CT_R());

            //ClassicAssert.IsTrue(p.RemoveRun(0));

            ClassicAssert.IsNotNull(p.Body);
            ClassicAssert.AreEqual(BodyElementType.PARAGRAPH, p.ElementType);
            ClassicAssert.AreEqual(BodyType.DOCUMENT, p.PartType);
        }

        [Test]
        public void TestSearchTextNotFound()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            ClassicAssert.IsNull(p.SearchText("test", new PositionInParagraph()));
            ClassicAssert.AreEqual("", p.Text);
        }

        [Test]
        public void TestSearchTextFound()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");

            IList<XWPFParagraph> ps = xml.Paragraphs;
            ClassicAssert.AreEqual(10, ps.Count);

            XWPFParagraph p = ps[(0)];

            TextSegment segment = p.SearchText("sample word document", new PositionInParagraph());
            ClassicAssert.IsNotNull(segment);

            ClassicAssert.AreEqual("sample word document", p.GetText(segment));

            ClassicAssert.IsTrue(p.RemoveRun(0));
        }

        [Test]
        public void TestFieldRuns()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("FldSimple.docx");
            IList<XWPFParagraph> ps = doc.Paragraphs;
            ClassicAssert.AreEqual(1, ps.Count);

            XWPFParagraph p = ps[0];
            ClassicAssert.AreEqual(1, p.Runs.Count);
            ClassicAssert.AreEqual(1, p.IRuns.Count);

            XWPFRun r = p.Runs[0];
            ClassicAssert.AreEqual(typeof(XWPFFieldRun), r.GetType());

            XWPFFieldRun fr = (XWPFFieldRun)r;
            ClassicAssert.AreEqual(" FILENAME   \\* MERGEFORMAT ", fr.FieldInstruction);
            ClassicAssert.AreEqual("FldSimple.docx", fr.Text);
            ClassicAssert.AreEqual("FldSimple.docx", p.Text);
        }

        [Test]
        public void TestRuns()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();
            IRunBody p1 = doc.CreateParagraph();

            CT_R run = new CT_R();
            XWPFRun r = new XWPFRun(run, p1);
            p.AddRun(r);
            p.AddRun(r);

            ClassicAssert.IsNotNull(p.GetRun(run));
            ClassicAssert.IsNull(p.GetRun(null));
        }

        [Test]
        public void TestAddingHyperlinks()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            
            XWPFParagraph p = doc.Paragraphs[0];

            ClassicAssert.AreEqual(2, p.Runs.Count);

            string rId = p.Part.GetPackagePart().AddExternalRelationship("https://www.google.com", XWPFRelation.HYPERLINK.Relation).Id;
            
            XWPFHyperlinkRun hr1 = p.CreateHyperlinkRun(rId);
            hr1.SetText("link1");
            ClassicAssert.AreEqual(3, p.Runs.Count);
            ClassicAssert.AreEqual(2, p.Runs.IndexOf(hr1));
            ClassicAssert.AreEqual(2, p.GetCTP().Items.IndexOf(hr1.GetCTHyperlink()));

            XWPFHyperlinkRun hr2 = p.InsertNewHyperlinkRun(1, rId);
            hr2.SetText("link2");
            ClassicAssert.AreEqual(4, p.Runs.Count);
            ClassicAssert.AreEqual(1, p.Runs.IndexOf(hr2));
            ClassicAssert.AreEqual(1, p.GetCTP().Items.IndexOf(hr2.GetCTHyperlink()));
        }

        [Test]
        public void Test58067()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("58067.docx");

            StringBuilder str = new StringBuilder();
            foreach (XWPFParagraph par in doc.Paragraphs)
            {
                str.Append(par.Text).Append("\n");
            }
            ClassicAssert.AreEqual("This is a test.\n\n\n\n3\n4\n5\n\n\n\nThis is a whole paragraph where one word is deleted.\n", str.ToString());
        }

        [Test]
        public void Test61787()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("61787.docx");

            StringBuilder str = new StringBuilder();
            foreach (XWPFParagraph par in doc.Paragraphs)
            {
                str.Append(par.Text).Append("\n");
            }
            String s = str.ToString();
            ClassicAssert.IsTrue(s.Trim().Length > 0, "Having text: \n" + s + "\nTrimmed lenght: " + s.Trim().Length);
        }

        [Test]
        public void Test61787_1()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("61787-1.docx");

            StringBuilder str = new StringBuilder();
            foreach (XWPFParagraph par in doc.Paragraphs)
            {
                str.Append(par.Text).Append("\n");
            }
            String s = str.ToString();
            ClassicAssert.IsFalse(s.Contains("This is another Test"));
        }

        [Test]
        public void Testpullrequest404()
        {
            XWPFDocument doc = new XWPFDocument();
            var paragraph = doc.CreateParagraph();
            paragraph.CreateRun().AppendText("abc");
            paragraph.CreateRun().AppendText("de");
            paragraph.CreateRun().AppendText("f");
            paragraph.CreateRun().AppendText("g");
            var result = paragraph.SearchText("cdefg", new PositionInParagraph());
            ClassicAssert.AreEqual(result.BeginRun, 0);
            ClassicAssert.AreEqual(result.EndRun, 3);
            ClassicAssert.AreEqual(result.BeginText, 0);
            ClassicAssert.AreEqual(result.EndText, 0);
            ClassicAssert.AreEqual(result.BeginChar, 2);
            ClassicAssert.AreEqual(result.EndChar, 0);
        }

        [Test]
        public void Testpullrequest404_1()
        {
            XWPFDocument doc = new XWPFDocument();
            var paragraph = doc.CreateParagraph();
            paragraph.CreateRun().AppendText("abc");
            paragraph.CreateRun().AppendText("de");
            paragraph.CreateRun().AppendText("fg");
            paragraph.CreateRun().AppendText("hi");
            var result = paragraph.SearchText("cdefg", new PositionInParagraph());
            ClassicAssert.AreEqual(result.BeginRun, 0);
            ClassicAssert.AreEqual(result.EndRun, 2);
            ClassicAssert.AreEqual(result.BeginText, 0);
            ClassicAssert.AreEqual(result.EndText, 0);
            ClassicAssert.AreEqual(result.BeginChar, 2);
            ClassicAssert.AreEqual(result.EndChar, 1);
        }

        /**
         * Tests for numbered lists
         * 
         * See also https://github.com/jimklo/apache-poi-sample/blob/master/src/main/java/com/sri/jklo/StyledDocument.java
         * for someone else trying a similar thing
         */
        [Test]
        public void TestNumberedLists()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("ComplexNumberedLists.docx");
            XWPFParagraph p;
        
            p = doc.GetParagraphArray(0);
            ClassicAssert.AreEqual("This is a document with numbered lists", p.Text);
            ClassicAssert.AreEqual(null, p.GetNumID());
            ClassicAssert.AreEqual(null, p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(1);
            ClassicAssert.AreEqual("Entry #1", p.Text);
            ClassicAssert.AreEqual("1", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(2);
            ClassicAssert.AreEqual("Entry #2, with children", p.Text);
            ClassicAssert.AreEqual("1", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(3);
            ClassicAssert.AreEqual("2-a", p.Text);
            ClassicAssert.AreEqual("1", p.GetNumID());
            ClassicAssert.AreEqual("1", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(4);
            ClassicAssert.AreEqual("2-b", p.Text);
            ClassicAssert.AreEqual("1", p.GetNumID());
            ClassicAssert.AreEqual("1", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(5);
            ClassicAssert.AreEqual("2-c", p.Text);
            ClassicAssert.AreEqual("1", p.GetNumID());
            ClassicAssert.AreEqual("1", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(6);
            ClassicAssert.AreEqual("Entry #3", p.Text);
            ClassicAssert.AreEqual("1", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(7);
            ClassicAssert.AreEqual("Entry #4", p.Text);
            ClassicAssert.AreEqual("1", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            // New list
            p = doc.GetParagraphArray(8);
            ClassicAssert.AreEqual("Restarted to 1 from 5", p.Text);
            ClassicAssert.AreEqual("2", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(9);
            ClassicAssert.AreEqual("Restarted @ 2", p.Text);
            ClassicAssert.AreEqual("2", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            p = doc.GetParagraphArray(10);
            ClassicAssert.AreEqual("Restarted @ 3", p.Text);
            ClassicAssert.AreEqual("2", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            // New list starting at 10
            p = doc.GetParagraphArray(11);
            ClassicAssert.AreEqual("Jump to new list at 10", p.Text);
            ClassicAssert.AreEqual("6", p.GetNumID());
            ClassicAssert.AreEqual("0", p.GetNumIlvl());
            // TODO Why isn't this seen as 10?
            ClassicAssert.AreEqual(null, p.GetNumStartOverride());
        
            // TODO Shouldn't we use XWPFNumbering or similar here?
            // TODO Make it easier to change
        }
    }
}

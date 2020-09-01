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
    using NUnit.Framework;
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

            Assert.AreEqual(Borders.Double, p.BorderTop);
            p.BorderTop = (Borders.Single);
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

            p.VerticalAlignment = (TextAlignment.BOTTOM);
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

            p.SetNumID("10");
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
            NPOI.OpenXmlFormats.Dml.CT_GraphicalObjectData god = r.GetCTR().GetDrawingArray(0).GetInlineArray(0).graphic.graphicData;
            //PicDocument pd = new PicDocumentImpl(null);
            //assertTrue(pd.isNil());
        }
        [Test]
        public void TestTika792()
        {
            //This test forces the loading of CTMoveBookmark and
            //CTMoveBookmarkImpl into ooxml-lite.
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Tika-792.docx");
            XWPFParagraph paragraph = doc.Paragraphs[(0)];
            Assert.AreEqual("s", paragraph.Text);
        }

        [Test]
        public void TestSettersGetters()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            //Assert.IsTrue(p.IsEmpty);
            Assert.IsFalse(p.RemoveRun(0));

            p.BorderTop = (/*setter*/Borders.BabyPacifier);
            p.BorderBetween = (/*setter*/Borders.BabyPacifier);
            p.BorderBottom = (/*setter*/Borders.BabyPacifier);

            Assert.IsNotNull(p.IRuns);
            Assert.AreEqual(0, p.IRuns.Count);
            Assert.IsFalse(p.IsEmpty);
            Assert.IsNull(p.StyleID);
            Assert.IsNull(p.Style);

            Assert.IsNull(p.GetNumID());
            p.SetNumID("12");
            Assert.AreEqual("12", p.GetNumID());
            p.SetNumID("13");
            Assert.AreEqual("13", p.GetNumID());

            Assert.IsNull(p.GetNumFmt());

            //Assert.IsNull(p.GetNumIlvl());

            Assert.AreEqual("", p.ParagraphText);
            Assert.AreEqual("", p.PictureText);
            Assert.AreEqual("", p.FootnoteText);

            p.BorderBetween = (/*setter*/Borders.None);
            Assert.AreEqual(Borders.None, p.BorderBetween);
            p.BorderBetween = (/*setter*/Borders.BasicBlackDashes);
            Assert.AreEqual(Borders.BasicBlackDashes, p.BorderBetween);

            p.BorderBottom = (/*setter*/Borders.None);
            Assert.AreEqual(Borders.None, p.BorderBottom);
            p.BorderBottom = (/*setter*/Borders.BabyPacifier);
            Assert.AreEqual(Borders.BabyPacifier, p.BorderBottom);

            p.BorderLeft = (/*setter*/Borders.None);
            Assert.AreEqual(Borders.None, p.BorderLeft);
            p.BorderLeft = (/*setter*/Borders.BasicWhiteSquares);
            Assert.AreEqual(Borders.BasicWhiteSquares, p.BorderLeft);

            p.BorderRight = (/*setter*/Borders.None);
            Assert.AreEqual(Borders.None, p.BorderRight);
            p.BorderRight = (/*setter*/Borders.BasicWhiteDashes);
            Assert.AreEqual(Borders.BasicWhiteDashes, p.BorderRight);

            p.BorderBottom = (/*setter*/Borders.None);
            Assert.AreEqual(Borders.None, p.BorderBottom);
            p.BorderBottom = (/*setter*/Borders.BasicWhiteDots);
            Assert.AreEqual(Borders.BasicWhiteDots, p.BorderBottom);

            Assert.IsFalse(p.IsPageBreak);
            p.IsPageBreak = (/*setter*/true);
            Assert.IsTrue(p.IsPageBreak);
            p.IsPageBreak = (/*setter*/false);
            Assert.IsFalse(p.IsPageBreak);

            Assert.AreEqual(-1, p.SpacingAfter);
            p.SpacingAfter = (/*setter*/12);
            Assert.AreEqual(12, p.SpacingAfter);

            Assert.AreEqual(-1, p.SpacingAfterLines);
            p.SpacingAfterLines = (/*setter*/14);
            Assert.AreEqual(14, p.SpacingAfterLines);

            Assert.AreEqual(-1, p.SpacingBefore);
            p.SpacingBefore = (/*setter*/16);
            Assert.AreEqual(16, p.SpacingBefore);

            Assert.AreEqual(-1, p.SpacingBeforeLines);
            p.SpacingBeforeLines = (/*setter*/18);
            Assert.AreEqual(18, p.SpacingBeforeLines);

            Assert.AreEqual(LineSpacingRule.AUTO, p.SpacingLineRule);
            p.SpacingLineRule = (/*setter*/LineSpacingRule.EXACT);
            Assert.AreEqual(LineSpacingRule.EXACT, p.SpacingLineRule);

            Assert.AreEqual(-1, p.IndentationLeft);
            p.IndentationLeft = (/*setter*/21);
            Assert.AreEqual(21, p.IndentationLeft);

            Assert.AreEqual(-1, p.IndentationRight);
            p.IndentationRight = (/*setter*/25);
            Assert.AreEqual(25, p.IndentationRight);

            Assert.AreEqual(-1, p.IndentationHanging);
            p.IndentationHanging = (/*setter*/25);
            Assert.AreEqual(25, p.IndentationHanging);

            Assert.AreEqual(-1, p.IndentationFirstLine);
            p.IndentationFirstLine = (/*setter*/25);
            Assert.AreEqual(25, p.IndentationFirstLine);

            Assert.IsFalse(p.IsWordWrap);
            p.IsWordWrap = (/*setter*/true);
            Assert.IsTrue(p.IsWordWrap);
            p.IsWordWrap = (/*setter*/false);
            Assert.IsFalse(p.IsWordWrap);

            Assert.IsNull(p.Style);
            p.Style = (/*setter*/"teststyle");
            Assert.AreEqual("teststyle", p.Style);

            p.AddRun(new CT_R());

            //Assert.IsTrue(p.RemoveRun(0));

            Assert.IsNotNull(p.Body);
            Assert.AreEqual(BodyElementType.PARAGRAPH, p.ElementType);
            Assert.AreEqual(BodyType.DOCUMENT, p.PartType);
        }

        [Test]
        public void TestSearchTextNotFound()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            Assert.IsNull(p.SearchText("test", new PositionInParagraph()));
            Assert.AreEqual("", p.Text);
        }

        [Test]
        public void TestSearchTextFound()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");

            IList<XWPFParagraph> ps = xml.Paragraphs;
            Assert.AreEqual(10, ps.Count);

            XWPFParagraph p = ps[(0)];

            TextSegment segment = p.SearchText("sample word document", new PositionInParagraph());
            Assert.IsNotNull(segment);

            Assert.AreEqual("sample word document", p.GetText(segment));

            Assert.IsTrue(p.RemoveRun(0));
        }

        [Test]
        public void TestFieldRuns()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("FldSimple.docx");
            IList<XWPFParagraph> ps = doc.Paragraphs;
            Assert.AreEqual(1, ps.Count);

            XWPFParagraph p = ps[0];
            Assert.AreEqual(1, p.Runs.Count);
            Assert.AreEqual(1, p.IRuns.Count);

            XWPFRun r = p.Runs[0];
            Assert.AreEqual(typeof(XWPFFieldRun), r.GetType());

            XWPFFieldRun fr = (XWPFFieldRun)r;
            Assert.AreEqual(" FILENAME   \\* MERGEFORMAT ", fr.FieldInstruction);
            Assert.AreEqual("FldSimple.docx", fr.Text);
            Assert.AreEqual("FldSimple.docx", p.Text);
        }

        [Test]
        public void TestRuns()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            CT_R run = new CT_R();
            XWPFRun r = new XWPFRun(run, doc.CreateParagraph());
            p.AddRun(r);
            p.AddRun(r);

            Assert.IsNotNull(p.GetRun(run));
            Assert.IsNull(p.GetRun(null));
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
            Assert.AreEqual("This is a test.\n\n\n\n3\n4\n5\n\n\n\nThis is a whole paragraph where one word is deleted.\n", str.ToString());
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
            Assert.IsTrue(s.Trim().Length > 0, "Having text: \n" + s + "\nTrimmed lenght: " + s.Trim().Length);
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
            Assert.IsFalse(s.Contains("This is another Test"));
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
            var result = paragraph.SearchText("cdefg",new PositionInParagraph());
            Assert.AreEqual(result.BeginRun, 0);
            Assert.AreEqual(result.EndRun, 3);
            Assert.AreEqual(result.BeginText, 0);
            Assert.AreEqual(result.EndText, 0);
            Assert.AreEqual(result.BeginChar, 2);
            Assert.AreEqual(result.EndChar, 0);
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
            var result = paragraph.SearchText("cdefg",new PositionInParagraph());
            Assert.AreEqual(result.BeginRun, 0);
            Assert.AreEqual(result.EndRun, 2);
            Assert.AreEqual(result.BeginText, 0);
            Assert.AreEqual(result.EndText, 0);
            Assert.AreEqual(result.BeginChar, 2);
            Assert.AreEqual(result.EndChar, 1);
        }
    }

}
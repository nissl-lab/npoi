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
    using CT_Blip = NPOI.OpenXmlFormats.Dml.CT_Blip;
    using CT_BlipFillProperties = NPOI.OpenXmlFormats.Dml.CT_BlipFillProperties;
    using CT_Picture = NPOI.OpenXmlFormats.Dml.Picture.CT_Picture;
    using NPOI.Util;
    using NPOI.WP.UserModel;
    using NPOI.XWPF.Model;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    /**
     * Tests for XWPF Run
     */
    [TestFixture]
    public class TestXWPFRun
    {

        public CT_R ctRun;
        public IRunBody p;
        [SetUp]
        public void SetUp()
        {
            XWPFDocument doc = new XWPFDocument();
            p = doc.CreateParagraph();

            this.ctRun = new CT_R();

        }

        [Test]
        public void TestSetGetText()
        {
            ctRun.AddNewT().Value = ("TEST STRING");
            ctRun.AddNewT().Value = ("TEST2 STRING");
            ctRun.AddNewT().Value = ("TEST3 STRING");

            Assert.AreEqual(3, ctRun.SizeOfTArray());
            XWPFRun run = new XWPFRun(ctRun, p);

            Assert.AreEqual("TEST2 STRING", run.GetText(1));

            run.SetText("NEW STRING", 0);
            Assert.AreEqual("NEW STRING", run.GetText(0));

            //Run.Text=("xxx",14);
            //Assert.Fail("Position wrong");
        }

        /*
         * bug 59208
         * Purpose: test all valid boolean-like values
         * exercise isCTOnOff(CTOnOff) through all valid permutations
         */
        [Ignore("stub testCTOnOff")]
        public void TestCTOnOff()
        {
            //CTRPr rpr = ctRun.addNewRPr();
            //CTOnOff bold = rpr.addNewB();
            //XWPFRun run = new XWPFRun(ctRun, p);

            //// True values: "true", "1", "on"
            //bold.setVal(STOnOff.TRUE);
            //assertEquals(true, run.isBold());

            //bold.setVal(STOnOff.X_1);
            //assertEquals(true, run.isBold());

            //bold.setVal(STOnOff.ON);
            //assertEquals(true, run.isBold());

            //// False values: "false", "0", "off"
            //bold.setVal(STOnOff.FALSE);
            //assertEquals(false, run.isBold());

            //bold.setVal(STOnOff.X_0);
            //assertEquals(false, run.isBold());

            //bold.setVal(STOnOff.OFF);
            //assertEquals(false, run.isBold());
        }

        [Test]
        public void TestSetGetBold()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewB().val = true;

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual(true, run.IsBold);

            run.IsBold = (false);
            // Implementation detail: POI natively prefers <w:b w:val="false"/>,
            // but should correctly read val="0" and val="off"
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, rpr.b.val);
        }

        [Test]
        public void TestSetGetItalic()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewI().val = true;

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual(true, run.IsItalic);

            run.IsItalic = false;
            Assert.AreEqual(false, rpr.i.val);
        }

        [Test]
        public void TestSetGetStrike()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewStrike().val = true;

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.IsTrue(run.IsStrikeThrough);

            run.IsStrikeThrough = false;
            Assert.AreEqual(false, rpr.strike.val);
        }

        [Test]
        public void TestSetGetUnderline()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewU().val = (ST_Underline.dash);

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual(UnderlinePatterns.Dash, run.Underline);

            run.Underline = UnderlinePatterns.None;
            Assert.AreEqual(ST_Underline.none, rpr.u.val);
        }


        [Test]
        public void TestSetGetVAlign()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewVertAlign().val = (ST_VerticalAlignRun.subscript);

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual(VerticalAlign.SUBSCRIPT, run.Subscript);

            run.Subscript = (VerticalAlign.BASELINE);
            Assert.AreEqual(ST_VerticalAlignRun.baseline, rpr.vertAlign.val);
        }


        [Test]
        public void TestSetGetFontFamily()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewRFonts().ascii = ("Times New Roman");

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual("Times New Roman", run.FontFamily);

            run.FontFamily = ("Verdana");
            Assert.AreEqual("Verdana", rpr.rFonts.ascii);
        }


        [Test]
        public void TestSetGetFontSize()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewSz().val = 14;

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual(7.0, run.FontSize);

            run.FontSize = 24;
            Assert.AreEqual(48, (int)rpr.sz.val);

            run.FontSize = 24.5;
            Assert.AreEqual(24.5, run.FontSize);
        }


        [Test]
        public void TestSetGetTextForegroundBackground()
        {
            CT_RPr rpr = ctRun.AddNewRPr();
            rpr.AddNewPosition().val = "4000";

            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual(4000, run.TextPosition);

            run.TextPosition = (2400);
            Assert.AreEqual(2400, int.Parse(rpr.position.val));
        }
        [Test]
        public void TestSetGetColor()
        {
            XWPFRun run = new XWPFRun(ctRun, p);
            run.SetColor("0F0F0F");
            String clr = run.GetColor();
            Assert.AreEqual("0F0F0F", clr);
        }
        [Test]
        public void TestAddCarriageReturn()
        {

            ctRun.AddNewT().Value = ("TEST STRING");
            ctRun.AddNewCr();
            ctRun.AddNewT().Value = ("TEST2 STRING");
            ctRun.AddNewCr();
            ctRun.AddNewT().Value = ("TEST3 STRING");
            Assert.AreEqual(2, ctRun.SizeOfCrArray());

            XWPFRun run = new XWPFRun(new CT_R(), p);
            run.AppendText("T1");
            run.AddCarriageReturn();
            run.AddCarriageReturn();
            run.AppendText("T2");
            run.AddCarriageReturn();
            Assert.AreEqual(3, run.GetCTR().GetCrList().Count);
            Assert.AreEqual("T1\n\nT2\n", run.ToString());

        }

        [Test]
        public void TestAddTabsAndLineBreaks()
        {
            ctRun.AddNewT().Value = "TEST STRING";
            ctRun.AddNewCr();
            ctRun.AddNewT().Value = "TEST2 STRING";
            ctRun.AddNewTab();
            ctRun.AddNewT().Value = "TEST3 STRING";
            Assert.AreEqual(1, ctRun.SizeOfCrArray());
            Assert.AreEqual(1, ctRun.SizeOfTabArray());

            XWPFRun run = new XWPFRun(new CT_R(), p);
            run.AppendText("T1");
            run.AddCarriageReturn();
            run.AppendText("T2");
            run.AddTab();
            run.AppendText("T3");
            Assert.AreEqual(1, run.GetCTR().GetCrList().Count);
            Assert.AreEqual(1, run.GetCTR().GetTabList().Count);

            Assert.AreEqual("T1\nT2\tT3", run.ToString());
        }


        [Test]
        public void TestAddPageBreak()
        {
            ctRun.AddNewT().Value = "TEST STRING";
            ctRun.AddNewBr();
            ctRun.AddNewT().Value = "TEST2 STRING";
            CT_Br breac = ctRun.AddNewBr();
            breac.clear = ST_BrClear.left;
            ctRun.AddNewT().Value = "TEST3 STRING";
            Assert.AreEqual(2, ctRun.SizeOfBrArray());

            XWPFRun run = new XWPFRun(new CT_R(), p);
            run.SetText("TEXT1");
            run.AddBreak();
            run.SetText("TEXT2");
            run.AddBreak(BreakType.TEXTWRAPPING);
            Assert.AreEqual(2, run.GetCTR().SizeOfBrArray());
        }

        /**
         * Test that on an existing document, we do the
         *  right thing with it
         * @throws IOException 
         */
        [Test]
        public void TestExisting()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
            XWPFParagraph p;
            XWPFRun run;


            // First paragraph is simple
            p = doc.GetParagraphArray(0);
            Assert.AreEqual("This is a test document.", p.Text);
            Assert.AreEqual(2, p.Runs.Count);

            run = p.Runs[0];
            Assert.AreEqual("This is a test document", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);

            run = p.Runs[1];
            Assert.AreEqual(".", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);


            // Next paragraph is all in one style, but a different one
            p = doc.GetParagraphArray(1);
            Assert.AreEqual("This bit is in bold and italic", p.Text);
            Assert.AreEqual(1, p.Runs.Count);

            run = p.Runs[0];
            Assert.AreEqual("This bit is in bold and italic", run.ToString());
            Assert.AreEqual(true, run.IsBold);
            Assert.AreEqual(true, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(true, run.GetCTR().rPr.IsSetB());
            Assert.AreEqual(false, run.GetCTR().rPr.b.IsSetVal());


            // Back to normal
            p = doc.GetParagraphArray(2);
            Assert.AreEqual("Back to normal", p.Text);
            Assert.AreEqual(1, p.Runs.Count);

            run = p.Runs[(0)];
            Assert.AreEqual("Back to normal", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);


            // Different styles in one paragraph
            p = doc.GetParagraphArray(3);
            Assert.AreEqual("This contains BOLD, ITALIC and BOTH, as well as RED and YELLOW text.", p.Text);
            Assert.AreEqual(11, p.Runs.Count);

            run = p.Runs[(0)];
            Assert.AreEqual("This contains ", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);

            run = p.Runs[(1)];
            Assert.AreEqual("BOLD", run.ToString());
            Assert.AreEqual(true, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);

            run = p.Runs[2];
            Assert.AreEqual(", ", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);

            run = p.Runs[(3)];
            Assert.AreEqual("ITALIC", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(true, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);

            run = p.Runs[(4)];
            Assert.AreEqual(" and ", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);

            run = p.Runs[(5)];
            Assert.AreEqual("BOTH", run.ToString());
            Assert.AreEqual(true, run.IsBold);
            Assert.AreEqual(true, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);

            run = p.Runs[(6)];
            Assert.AreEqual(", as well as ", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);

            run = p.Runs[(7)];
            Assert.AreEqual("RED", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);

            run = p.Runs[(8)];
            Assert.AreEqual(" and ", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);

            run = p.Runs[(9)];
            Assert.AreEqual("YELLOW", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);

            run = p.Runs[(10)];
            Assert.AreEqual(" text.", run.ToString());
            Assert.AreEqual(false, run.IsBold);
            Assert.AreEqual(false, run.IsItalic);
            Assert.IsFalse(run.IsStrikeThrough);
            Assert.AreEqual(null, run.GetCTR().rPr);
        }

        [Test]
        public void TestPictureInHeader()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("headerPic.docx");
            XWPFHeaderFooterPolicy policy = sampleDoc.GetHeaderFooterPolicy();

            XWPFHeader header = policy.GetDefaultHeader();

            int count = 0;

            foreach (XWPFParagraph p in header.Paragraphs)
            {
                foreach (XWPFRun r in p.Runs)
                {
                    List<XWPFPicture> pictures = r.GetEmbeddedPictures();

                    foreach (XWPFPicture pic in pictures)
                    {
                        Assert.IsNotNull(pic.GetPictureData());
                        Assert.AreEqual("DOZOR", pic.GetDescription());
                    }

                    count += pictures.Count;
                }
            }

            Assert.AreEqual(1, count);
        }

        [Test]
        public void testSetGetHighlight()
        {
            XWPFRun run = new XWPFRun(ctRun, p);
            Assert.AreEqual(false, run.IsHighlighted);

            // TODO Do this using XWPFRun methods
            run.GetCTR().AddNewRPr().AddNewHighlight().val = (ST_HighlightColor.none);
            Assert.AreEqual(false, run.IsHighlighted);
            run.GetCTR().rPr.highlight.val = (ST_HighlightColor.cyan);
            Assert.AreEqual(true, run.IsHighlighted);
            run.GetCTR().rPr.highlight.val = (ST_HighlightColor.none);
            Assert.AreEqual(false, run.IsHighlighted);
        }

        [Test]
        public void TestAddPicture()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
            XWPFParagraph p = doc.GetParagraphArray(2);
            XWPFRun r = p.Runs[0];

            Assert.AreEqual(0, doc.AllPictures.Count);
            Assert.AreEqual(0, r.GetEmbeddedPictures().Count);

            r.AddPicture(new MemoryStream(new byte[0]), (int)PictureType.JPEG, "test.jpg", 21, 32);

            Assert.AreEqual(1, doc.AllPictures.Count);
            Assert.AreEqual(1, r.GetEmbeddedPictures().Count);
        }

        /**
         * Bugzilla #58237 - Unable to add image to word document header
         *
         * @throws Exception
         */
        [Test]
        public void TestAddPictureInHeader()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
            XWPFHeader hdr = doc.CreateHeader(HeaderFooterType.DEFAULT);
            XWPFParagraph p = hdr.CreateParagraph();
            XWPFRun r = p.CreateRun();

            Assert.AreEqual(0, hdr.AllPictures.Count);
            Assert.AreEqual(0, r.GetEmbeddedPictures().Count);

            r.AddPicture(new ByteArrayInputStream(new byte[0]), (int)PictureType.JPEG, "test.jpg", 21, 32);

            Assert.AreEqual(1, hdr.AllPictures.Count);
            Assert.AreEqual(1, r.GetEmbeddedPictures().Count);

            XWPFPicture pic = r.GetEmbeddedPictures()[0];
            CT_Picture ctPic = pic.GetCTPicture();
            CT_BlipFillProperties ctBlipFill = ctPic.blipFill;

            Assert.IsNotNull(ctBlipFill);

            CT_Blip ctBlip = ctBlipFill.blip;

            Assert.IsNotNull(ctBlip);
            Assert.AreEqual("rId1", ctBlip.embed);

            XWPFDocument docBack = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            XWPFHeader hdrBack = docBack.GetHeaderArray(0);
            XWPFParagraph pBack = hdrBack.GetParagraphArray(0);
            XWPFRun rBack = pBack.Runs[0];

            Assert.AreEqual(1, hdrBack.AllPictures.Count);
            Assert.AreEqual(1, rBack.GetEmbeddedPictures().Count);
        }

        /**
         * Bugzilla #52288 - setting the font family on the
         *  run mustn't NPE
         */
        [Test]
        public void TestSetFontFamily_52288()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("52288.docx");
            IEnumerator<XWPFParagraph> paragraphs = doc.Paragraphs.GetEnumerator();
            while (paragraphs.MoveNext())
            {
                XWPFParagraph paragraph = paragraphs.Current;
                foreach (XWPFRun run in paragraph.Runs)
                {
                    if (run != null)
                    {
                        String text = run.GetText(0);
                        if (text != null)
                        {
                            run.FontFamily = ("Times New Roman");
                        }
                    }
                }
            }
        }

        [Test]
        public void TestBug55476()
        {
            byte[] image = XWPFTestDataSamples.GetImage("abstract1.jpg");
            XWPFDocument document = new XWPFDocument();
            document.CreateParagraph().CreateRun().AddPicture(
                    new MemoryStream(image), (int)PictureType.JPEG, "test.jpg", Units.ToEMU(300), Units.ToEMU(100));
            XWPFDocument docBack = XWPFTestDataSamples.WriteOutAndReadBack(document);
            List<XWPFPicture> pictures = docBack.GetParagraphArray(0).Runs[0].GetEmbeddedPictures();
            Assert.AreEqual(1, pictures.Count);
            docBack.Close();
            /*OutputStream stream = new FileOutputStream("c:\\temp\\55476.docx");
            try {
                document.write(stream);
            } finally {
                stream.close();
            }*/
            document.Close();
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestBug58922()
        {
            XWPFDocument document = new XWPFDocument();
            XWPFRun run = document.CreateParagraph().CreateRun();
            Assert.AreEqual(-1, run.FontSize);
            run.FontSize = 10;
            Assert.AreEqual(10, run.FontSize);
            run.FontSize = short.MaxValue - 1;
            Assert.AreEqual(short.MaxValue - 1, run.FontSize);
            run.FontSize = short.MaxValue;
            Assert.AreEqual(short.MaxValue, run.FontSize);
            run.FontSize = short.MaxValue + 1;
            Assert.AreEqual(short.MaxValue + 1, run.FontSize);
            run.FontSize = int.MaxValue - 1;
            Assert.AreEqual(int.MaxValue - 1, run.FontSize);
            run.FontSize = int.MaxValue;
            Assert.AreEqual(int.MaxValue, run.FontSize);
            run.FontSize = -1;
            Assert.AreEqual(-1, run.FontSize);
            Assert.AreEqual(-1, run.TextPosition);
            run.TextPosition = 10;
            Assert.AreEqual(10, run.TextPosition);
            run.TextPosition = short.MaxValue - 1;
            Assert.AreEqual(short.MaxValue - 1, run.TextPosition);
            run.TextPosition = short.MaxValue;
            Assert.AreEqual(short.MaxValue, run.TextPosition);
            run.TextPosition = short.MaxValue + 1;
            Assert.AreEqual(short.MaxValue + 1, run.TextPosition);
            run.TextPosition = short.MaxValue + 1;
            Assert.AreEqual(short.MaxValue + 1, run.TextPosition);
            run.TextPosition = int.MaxValue - 1;
            Assert.AreEqual(int.MaxValue - 1, run.TextPosition);
            run.TextPosition = int.MaxValue;
            Assert.AreEqual(int.MaxValue, run.TextPosition);
            run.TextPosition = -1;
            Assert.AreEqual(-1, run.TextPosition);
        }

        [Test]
        public void TestWhitespace()
        {
            String[]
            text = new String[] {
                "  The quick brown fox",
                "\t\tjumped over the lazy dog"
            };
            MemoryStream bos = new MemoryStream();
            XWPFDocument doc = new XWPFDocument();
            foreach (String s in text)
            {
                XWPFParagraph p1 = doc.CreateParagraph();
                XWPFRun r1 = p1.CreateRun();
                r1.SetText(s);
            }
            doc.Write(bos);

            MemoryStream bis = new MemoryStream(bos.ToArray());
            var doc2 = new XWPFDocument(bis);

            var paragraphs = doc2.Paragraphs;
            Assert.AreEqual(2, paragraphs.Count);
            for (int i = 0; i < text.Length; i++)
            {
                XWPFParagraph p1 = paragraphs[i];
                String expected = text[i];
                Assert.AreEqual(expected, p1.Text);
                CT_P ctp = p1.GetCTP();
                CT_R ctr = ctp.GetRArray(0);
                CT_Text ctText = ctr.GetTArray(0);
                // if text has leading whitespace then expect xml-fragment to have xml:space="preserve" set
                // <xml-fragment xml:space="preserve" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                bool isWhitespace = Character.isWhitespace(expected[0]);
                Assert.AreEqual(isWhitespace, ctText.space == "preserve");
            }
        }
        [Test]
        public void TestSetStyleId()
        {
            XWPFDocument document = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");

            XWPFRun run = document.CreateParagraph().CreateRun();

            String styleId = "bolditalic";
            run.SetStyle(styleId);
            String candStyleId = run.GetCTR().rPr.rStyle.val;
            Assert.IsNotNull(candStyleId, "Expected to find a run style ID");
            Assert.AreEqual(styleId, candStyleId);

            Assert.AreEqual(styleId, run.GetStyle());

            document.Close();
        }
    }
}

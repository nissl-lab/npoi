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
using NPOI.XWPF.Model;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;

/**
 * Tests for XWPF Run
 */
    [TestClass]
    public class TestXWPFRun 
{

    public CTR ctRun;
    public XWPFParagraph p;

    protected void SetUp() {
        XWPFDocument doc = new XWPFDocument();
        p = doc.CreateParagraph();

        this.ctRun = CTR.Factory.NewInstance();
        
    }

        [TestMethod]
    public void TestSetGetText(){
	ctRun.AddNewT().StringValue=("TEST STRING");	
	ctRun.AddNewT().StringValue=("TEST2 STRING");	
	ctRun.AddNewT().StringValue=("TEST3 STRING");
	
	Assert.AreEqual(3,ctRun.SizeOfTArray());
	XWPFRun run = new XWPFRun(ctRun, p);
	
	Assert.AreEqual("TEST2 STRING",Run.GetText(1));
	
	Run.Text=("NEW STRING",0);
	Assert.AreEqual("NEW STRING",Run.GetText(0));
	
	//Run.Text=("xxx",14);
	//Assert.Fail("Position wrong");
    }
  
        [TestMethod]
    public void TestSetGetBold(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewB().Val=(STOnOff.TRUE);

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual(true, Run.IsBold());

        Run.Bold=(false);
        Assert.AreEqual(STOnOff.FALSE, rpr.B.Val);
    }

        [TestMethod]
    public void TestSetGetItalic(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewI().Val=(STOnOff.TRUE);

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual(true, Run.IsItalic());

        Run.Italic=(false);
        Assert.AreEqual(STOnOff.FALSE, rpr.I.Val);
    }

        [TestMethod]
    public void TestSetGetStrike(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewStrike().Val=(STOnOff.TRUE);

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual(true, Run.IsStrike());

        Run.Strike=(false);
        Assert.AreEqual(STOnOff.FALSE, rpr.Strike.Val);
    }

        [TestMethod]
    public void TestSetGetUnderline(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewU().Val=(STUnderline.DASH);

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual(UnderlinePatterns.DASH.Value, Run.Underline
                .Value);

        Run.Underline=(UnderlinePatterns.NONE);
        Assert.AreEqual(STUnderline.NONE.IntValue(), rpr.U.Val
                .intValue());
    }


        [TestMethod]
    public void TestSetGetVAlign(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewVertAlign().Val=(STVerticalAlignRun.SUBSCRIPT);

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual(VerticalAlign.SUBSCRIPT, Run.Subscript);

        Run.Subscript=(VerticalAlign.BASELINE);
        Assert.AreEqual(STVerticalAlignRun.BASELINE, rpr.VertAlign.Val);
    }


        [TestMethod]
    public void TestSetGetFontFamily(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewRFonts().Ascii=("Times New Roman");

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual("Times New Roman", Run.FontFamily);

        Run.FontFamily=("Verdana");
        Assert.AreEqual("Verdana", rpr.RFonts.Ascii);
    }


        [TestMethod]
    public void TestSetGetFontSize(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewSz().Val=(new Bigint("14"));

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual(7, Run.FontSize);

        Run.FontSize=(24);
        Assert.AreEqual(48, rpr.Sz.Val.LongValue());
    }


        [TestMethod]
    public void TestSetGetTextForegroundBackground(){
        CTRPr rpr = ctRun.AddNewRPr();
        rpr.AddNewPosition().Val=(new Bigint("4000"));

        XWPFRun run = new XWPFRun(ctRun, p);
        Assert.AreEqual(4000, Run.TextPosition);

        Run.TextPosition=(2400);
        Assert.AreEqual(2400, rpr.Position.Val.LongValue());
    }

        [TestMethod]
    public void TestAddCarriageReturn(){
	
	ctRun.AddNewT().StringValue=("TEST STRING");
	ctRun.AddNewCr();
	ctRun.AddNewT().StringValue=("TEST2 STRING");
	ctRun.AddNewCr();
	ctRun.AddNewT().StringValue=("TEST3 STRING");
        Assert.AreEqual(2, ctRun.SizeOfCrArray());
        
        XWPFRun run = new XWPFRun(CTR.Factory.NewInstance(), p);
        Run.Text=("T1");
        Run.AddCarriageReturn();
        Run.AddCarriageReturn();
        Run.Text=("T2");
        Run.AddCarriageReturn();
        Assert.AreEqual(3, Run.CTR.CrList.Size());
        
    }

        [TestMethod]
    public void TestAddPageBreak(){
	ctRun.AddNewT().StringValue=("TEST STRING");
	ctRun.AddNewBr();
	ctRun.AddNewT().StringValue=("TEST2 STRING");
	CTBr breac=ctRun.AddNewBr();
	breac.Clear=(STBrClear.LEFT);
	ctRun.AddNewT().StringValue=("TEST3 STRING");
        Assert.AreEqual(2, ctRun.SizeOfBrArray());
        
        XWPFRun run = new XWPFRun(CTR.Factory.NewInstance(), p);
        Run.Text=("TEXT1");
        Run.AddBreak();
        Run.Text=("TEXT2");
        Run.AddBreak(BreakType.TEXT_WRAPPING);
        Assert.AreEqual(2, Run.CTR.SizeOfBrArray());
    }

    /**
     * Test that on an existing document, we do the
     *  right thing with it
     * @throws IOException 
     */
        [TestMethod]
    public void TestExisting(){
       XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
       XWPFParagraph p;
       XWPFRun Run;
       
       
       // First paragraph is simple
       p = doc.GetParagraphArray(0);
       Assert.AreEqual("This is a test document.", p.Text);
       Assert.AreEqual(2, p.Runs.Size());
       
       run = p.Runs.Get(0);
       Assert.AreEqual("This is a test document", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       run = p.Runs.Get(1);
       Assert.AreEqual(".", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       
       // Next paragraph is all in one style, but a different one
       p = doc.GetParagraphArray(1);
       Assert.AreEqual("This bit is in bold and italic", p.Text);
       Assert.AreEqual(1, p.Runs.Size());
       
       run = p.Runs.Get(0);
       Assert.AreEqual("This bit is in bold and italic", Run.ToString());
       Assert.AreEqual(true, Run.IsBold());
       Assert.AreEqual(true, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(true, Run.CTR.RPr.IsSetB());
       Assert.AreEqual(false, Run.CTR.RPr.B.IsSetVal());
       
       
       // Back to normal
       p = doc.GetParagraphArray(2);
       Assert.AreEqual("Back to normal", p.Text);
       Assert.AreEqual(1, p.Runs.Size());
       
       run = p.Runs.Get(0);
       Assert.AreEqual("Back to normal", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       
       // Different styles in one paragraph
       p = doc.GetParagraphArray(3);
       Assert.AreEqual("This Contains BOLD, ITALIC and BOTH, as well as RED and YELLOW text.", p.Text);
       Assert.AreEqual(11, p.Runs.Size());
       
       run = p.Runs.Get(0);
       Assert.AreEqual("This Contains ", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       run = p.Runs.Get(1);
       Assert.AreEqual("BOLD", Run.ToString());
       Assert.AreEqual(true, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       
       run = p.Runs.Get(2);
       Assert.AreEqual(", ", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       run = p.Runs.Get(3);
       Assert.AreEqual("ITALIC", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(true, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       
       run = p.Runs.Get(4);
       Assert.AreEqual(" and ", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       run = p.Runs.Get(5);
       Assert.AreEqual("BOTH", Run.ToString());
       Assert.AreEqual(true, Run.IsBold());
       Assert.AreEqual(true, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       
       run = p.Runs.Get(6);
       Assert.AreEqual(", as well as ", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       run = p.Runs.Get(7);
       Assert.AreEqual("RED", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       
       run = p.Runs.Get(8);
       Assert.AreEqual(" and ", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
       
       run = p.Runs.Get(9);
       Assert.AreEqual("YELLOW", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       
       run = p.Runs.Get(10);
       Assert.AreEqual(" text.", Run.ToString());
       Assert.AreEqual(false, Run.IsBold());
       Assert.AreEqual(false, Run.IsItalic());
       Assert.AreEqual(false, Run.IsStrike());
       Assert.AreEqual(null, Run.CTR.RPr);
    }

        [TestMethod]
    public void TestPictureInHeader(){
        XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("headerPic.docx");
        XWPFHeaderFooterPolicy policy = sampleDoc.HeaderFooterPolicy;

        XWPFHeader header = policy.DefaultHeader;

        int count = 0;

        foreach (XWPFParagraph p in header.Paragraphs) {
            foreach (XWPFRun r in p.Runs) {
                List<XWPFPicture> pictures = r.EmbeddedPictures;

                foreach (XWPFPicture pic in pictures) {
                    Assert.IsNotNull(pic.PictureData);
                    Assert.AreEqual("DOZOR", pic.Description);
                }

                count+= pictures.Size();
            }
        }

        Assert.AreEqual(1, count);
    }
    
        [TestMethod]
    public void TestAddPicture(){
       XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
       XWPFParagraph p = doc.GetParagraphArray(2);
       XWPFRun r = p.Runs.Get(0);
       
       Assert.AreEqual(0, doc.AllPictures.Size());
       Assert.AreEqual(0, r.EmbeddedPictures.Size());
       
       r.AddPicture(new MemoryStream(new byte[0]), Document.PICTURE_TYPE_JPEG, "test.jpg", 21, 32);
       
       Assert.AreEqual(1, doc.AllPictures.Size());
       Assert.AreEqual(1, r.EmbeddedPictures.Size());
    }
}


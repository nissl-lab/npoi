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

    [TestClass]
    public class TestXWPFHeader 
{

	    [TestMethod]
    public void TestSimpleHeader(){
		XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("headerFooter.docx");

		XWPFHeaderFooterPolicy policy = sampleDoc.HeaderFooterPolicy;

		XWPFHeader header = policy.DefaultHeader;
		XWPFFooter footer = policy.DefaultFooter;
		Assert.IsNotNull(header);
		Assert.IsNotNull(footer);
	}

        [TestMethod]
    public void TestImageInHeader(){
        XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("headerPic.docx");

        XWPFHeaderFooterPolicy policy = sampleDoc.HeaderFooterPolicy;

        XWPFHeader header = policy.DefaultHeader;

        Assert.IsNotNull(header.Relations);
        Assert.AreEqual(1, header.Relations.Size());
    }

	    [TestMethod]
    public void TestSetHeader(){
		XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
		// no header is Set (yet)
		XWPFHeaderFooterPolicy policy = sampleDoc.HeaderFooterPolicy;
		Assert.IsNull(policy.DefaultHeader);
		Assert.IsNull(policy.FirstPageHeader);
		Assert.IsNull(policy.DefaultFooter);

		CTP ctP1 = CTP.Factory.NewInstance();
		CTR ctR1 = ctP1.AddNewR();
		CTText t = ctR1.AddNewT();
		t.StringValue=("Paragraph in header");

		// Commented MB 23 May 2010
		//CTP ctP2 = CTP.Factory.NewInstance();
		//CTR ctR2 = ctP2.AddNewR();
		//CTText t2 = ctR2.AddNewT();
		//t2.StringValue=("Second paragraph.. for footer");
		
		// Create two paragraphs for insertion into the footer.
		// Previously only one was inserted MB 23 May 2010
		CTP ctP2 = CTP.Factory.NewInstance();
		CTR ctR2 = ctP2.AddNewR();
		CTText t2 = ctR2.AddNewT();
		t2.StringValue=("First paragraph for the footer");
		
		CTP ctP3 = CTP.Factory.NewInstance();
		CTR ctR3 = ctP3.AddNewR();
		CTText t3 = ctR3.AddNewT();
		t3.StringValue=("Second paragraph for the footer");

		XWPFParagraph p1 = new XWPFParagraph(ctP1, sampleDoc);
		XWPFParagraph[] pars = new XWPFParagraph[1];
		pars[0] = p1;

		XWPFParagraph p2 = new XWPFParagraph(ctP2, sampleDoc);
		XWPFParagraph p3 = new XWPFParagraph(ctP3, sampleDoc);
		XWPFParagraph[] pars2 = new XWPFParagraph[2];
		pars2[0] = p2;
		pars2[1] = p3;

		// Set headers
		policy.CreateHeader(policy.DEFAULT, pars);
		policy.CreateHeader(policy.FIRST);
		// Set a default footer and capture the returned XWPFFooter object.
		XWPFFooter footer = policy.CreateFooter(policy.DEFAULT, pars2);

		// Ensure the headers and footer were Set correctly....
		Assert.IsNotNull(policy.DefaultHeader);
		Assert.IsNotNull(policy.FirstPageHeader);
		Assert.IsNotNull(policy.DefaultFooter);
		// ....and that the footer object captured above Contains two
		// paragraphs of text.
		Assert.AreEqual(2, footer.Paragraphs.Size());
		
		// As an Additional Check, recover the defauls footer and
		// make sure that it Contains two paragraphs of text and that
		// both do hold what is expected.
		footer = policy.DefaultFooter;
		
		XWPFParagraph[] paras = new XWPFParagraph[footer.Paragraphs.Size()];
		int i=0;
		foreach(XWPFParagraph p in footer.Paragraphs) {
		   paras[i++] = p;
		}
		
		Assert.AreEqual(2, paras.Length);
		Assert.AreEqual("First paragraph for the footer", paras[0].Text);
		Assert.AreEqual("Second paragraph for the footer", paras[1].Text);
	}

	    [TestMethod]
    public void TestSetWatermark(){
		XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
		// no header is Set (yet)
		XWPFHeaderFooterPolicy policy = sampleDoc.HeaderFooterPolicy;
		Assert.IsNull(policy.DefaultHeader);
		Assert.IsNull(policy.FirstPageHeader);
		Assert.IsNull(policy.DefaultFooter);

		policy.CreateWatermark("DRAFT");

		Assert.IsNotNull(policy.DefaultHeader);
		Assert.IsNotNull(policy.FirstPageHeader);
		Assert.IsNotNull(policy.EvenPageHeader);
	}
	
	    [TestMethod]
    public void TestAddPictureData(){
	    
	}
	
	    [TestMethod]
    public void TestGetAllPictures(){
	    
	}
	
	    [TestMethod]
    public void TestGetAllPackagePictures(){
	    
	}
	
	    [TestMethod]
    public void TestGetPictureDataById(){
	    
	}
}


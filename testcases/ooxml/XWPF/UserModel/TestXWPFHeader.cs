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
    using NPOI.XWPF.Model;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using System.Collections.Generic;

    [TestFixture]
    public class TestXWPFHeader
    {

        [Test]
        public void TestSimpleHeader()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("headerFooter.docx");

            XWPFHeaderFooterPolicy policy = sampleDoc.GetHeaderFooterPolicy();

            XWPFHeader header = policy.GetDefaultHeader();
            XWPFFooter footer = policy.GetDefaultFooter();
            ClassicAssert.IsNotNull(header);
            ClassicAssert.IsNotNull(footer);
        }

        [Test]
        public void TestImageInHeader()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("headerPic.docx");

            XWPFHeaderFooterPolicy policy = sampleDoc.GetHeaderFooterPolicy();

            XWPFHeader header = policy.GetDefaultHeader();

            ClassicAssert.IsNotNull(header.GetRelations());
            ClassicAssert.AreEqual(1, header.GetRelations().Count);
        }

        [Test]
        public void TestSetHeader()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
            // no header is Set (yet)
            XWPFHeaderFooterPolicy policy = sampleDoc.GetHeaderFooterPolicy();
            ClassicAssert.IsNull(policy.GetDefaultHeader());
            ClassicAssert.IsNull(policy.GetFirstPageHeader());
            ClassicAssert.IsNull(policy.GetDefaultFooter());
            ClassicAssert.IsNull(policy.GetFirstPageFooter());

            CT_P ctP1 = new CT_P();
            CT_R ctR1 = ctP1.AddNewR();
            CT_Text t = ctR1.AddNewT();
            String tText = "Paragraph in header";
            t.Value = tText;

            // Commented MB 23 May 2010
            //CTP ctP2 = CTP.Factory.NewInstance();
            //CTR ctR2 = ctP2.AddNewR();
            //CTText t2 = ctR2.AddNewT();
            //t2.StringValue=("Second paragraph.. for footer");

            // Create two paragraphs for insertion into the footer.
            // Previously only one was inserted MB 23 May 2010
            CT_P ctP2 = new CT_P();
            CT_R ctR2 = ctP2.AddNewR();
            CT_Text t2 = ctR2.AddNewT();
            t2.Value = ("First paragraph for the footer");

            CT_P ctP3 = new CT_P();
            CT_R ctR3 = ctP3.AddNewR();
            CT_Text t3 = ctR3.AddNewT();
            t3.Value = ("Second paragraph for the footer");

            XWPFParagraph p1 = new XWPFParagraph(ctP1, sampleDoc);
            XWPFParagraph[] pars = new XWPFParagraph[1];
            pars[0] = p1;

            XWPFParagraph p2 = new XWPFParagraph(ctP2, sampleDoc);
            XWPFParagraph p3 = new XWPFParagraph(ctP3, sampleDoc);
            XWPFParagraph[] pars2 = new XWPFParagraph[2];
            pars2[0] = p2;
            pars2[1] = p3;

            // Set headers
            XWPFHeader headerD = policy.CreateHeader(XWPFHeaderFooterPolicy.DEFAULT, pars);
            XWPFHeader headerF = policy.CreateHeader(XWPFHeaderFooterPolicy.FIRST);
            // Set a default footer and capture the returned XWPFFooter object.
            XWPFFooter footerD = policy.CreateFooter(XWPFHeaderFooterPolicy.DEFAULT, pars2);
            XWPFFooter footerF = policy.CreateFooter(XWPFHeaderFooterPolicy.FIRST);

            // Ensure the headers and footer were Set correctly....
            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetFirstPageHeader());
            ClassicAssert.IsNotNull(policy.GetDefaultFooter());
            ClassicAssert.IsNotNull(policy.GetFirstPageFooter());
            // ....and that the footer object captured above Contains two
            // paragraphs of text.
            ClassicAssert.AreEqual(2, footerD.Paragraphs.Count);
            ClassicAssert.AreEqual(0, footerF.Paragraphs.Count);

            // Check the header created with the paragraph got them, and the one
            // created without got none
            ClassicAssert.AreEqual(1, headerD.Paragraphs.Count);
            ClassicAssert.AreEqual(tText, headerD.Paragraphs[0].Text);

            ClassicAssert.AreEqual(0, headerF.Paragraphs.Count);

            // As an Additional Check, recover the defauls footer and
            // make sure that it Contains two paragraphs of text and that
            // both do hold what is expected.
            footerD = policy.GetDefaultFooter();
            XWPFParagraph[] paras = new List<XWPFParagraph>(footerD.Paragraphs).ToArray();

            ClassicAssert.AreEqual(2, paras.Length);
            ClassicAssert.AreEqual("First paragraph for the footer", paras[0].Text);
            ClassicAssert.AreEqual("Second paragraph for the footer", paras[1].Text);


            // Add some text to the empty header
            String fText1 = "New Text!";
            String fText2 = "More Text!";
            headerF.CreateParagraph().InsertNewRun(0).SetText(fText1);
            headerF.CreateParagraph().InsertNewRun(0).SetText(fText2);
            //headerF.Paragraphs[0].InsertNewRun(0).SetText(fText1);

            // Check it
            ClassicAssert.AreEqual(tText, headerD.Paragraphs[0].Text);
            ClassicAssert.AreEqual(fText1, headerF.Paragraphs[0].Text);
            ClassicAssert.AreEqual(fText2, headerF.Paragraphs[1].Text);


            // Save, re-open, ensure it's all still there
            XWPFDocument reopened = XWPFTestDataSamples.WriteOutAndReadBack(sampleDoc);
            policy = reopened.GetHeaderFooterPolicy();
            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetFirstPageHeader());
            ClassicAssert.IsNull(policy.GetEvenPageHeader());
            ClassicAssert.IsNotNull(policy.GetDefaultFooter());
            ClassicAssert.IsNotNull(policy.GetFirstPageFooter());
            ClassicAssert.IsNull(policy.GetEvenPageFooter());

            // Check the new headers still have their text
            headerD = policy.GetDefaultHeader();
            headerF = policy.GetFirstPageHeader();
            ClassicAssert.AreEqual(tText, headerD.Paragraphs[0].Text);
            ClassicAssert.AreEqual(fText1, headerF.Paragraphs[0].Text);
            ClassicAssert.AreEqual(fText2, headerF.Paragraphs[1].Text);

            // Check the new footers have their new text too
            footerD = policy.GetDefaultFooter();
            paras = new List<XWPFParagraph>(footerD.Paragraphs).ToArray();
            footerF = policy.GetFirstPageFooter();

            ClassicAssert.AreEqual(2, paras.Length);
            ClassicAssert.AreEqual("First paragraph for the footer", paras[0].Text);
            ClassicAssert.AreEqual("Second paragraph for the footer", paras[1].Text);
            ClassicAssert.AreEqual(1, footerF.Paragraphs.Count);
        }

        [Test]
        public void TestSetWatermark()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");

            // no header is Set (yet)
            XWPFHeaderFooterPolicy policy = sampleDoc.GetHeaderFooterPolicy();
            ClassicAssert.IsNull(policy.GetDefaultHeader());
            ClassicAssert.IsNull(policy.GetFirstPageHeader());
            ClassicAssert.IsNull(policy.GetDefaultFooter());

            policy.CreateWatermark("DRAFT");

            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetFirstPageHeader());
            ClassicAssert.IsNotNull(policy.GetEvenPageHeader());

            // Re-open, and check
            XWPFDocument reopened = XWPFTestDataSamples.WriteOutAndReadBack(sampleDoc);
            policy = reopened.GetHeaderFooterPolicy();

            ClassicAssert.IsNotNull(policy.GetDefaultHeader());
            ClassicAssert.IsNotNull(policy.GetFirstPageHeader());
            ClassicAssert.IsNotNull(policy.GetEvenPageHeader());
        }

        [Test]
        public void TestAddPictureData()
        {
            // TODO
        }

        [Test]
        public void TestGetAllPictures()
        {
            // TODO
        }

        [Test]
        public void TestGetAllPackagePictures()
        {
            // TODO
        }

        [Test]
        public void TestGetPictureDataById()
        {
            // TODO
        }

        [Test]
        public void Bug60293()
        {
            //test handling of non-standard header/footer options
            XWPFDocument xwpf = XWPFTestDataSamples.OpenSampleDocument("60293.docx");
            ClassicAssert.AreEqual(3, xwpf.HeaderList.Count);
        }
    }
}

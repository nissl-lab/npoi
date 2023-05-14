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
    using System;
    using NUnit.Framework;
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using TestCases.XWPF;
    using NPOI.XWPF.UserModel;

    [TestFixture]
    public class TestXWPFStyles
    {

        //	protected void SetUp()  {
        //		super.Up=();
        //	}

        [Test]
        public void TestGetUsedStyles()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("Styles.docx");
            List<XWPFStyle> testUsedStyleList = new List<XWPFStyle>();
            XWPFStyles styles = sampleDoc.GetStyles();
            XWPFStyle style = styles.GetStyle("berschrift1");
            testUsedStyleList.Add(style);
            testUsedStyleList.Add(styles.GetStyle("Standard"));
            testUsedStyleList.Add(styles.GetStyle("berschrift1Zchn"));
            testUsedStyleList.Add(styles.GetStyle("Absatz-Standardschriftart"));
            style.HasSameName(style);

            List<XWPFStyle> usedStyleList = styles.GetUsedStyleList(style);
            
            //Assert.AreEqual(usedStyleList, testUsedStyleList);
            Assert.AreEqual(usedStyleList.Count, testUsedStyleList.Count);
            for (int i = 0; i < usedStyleList.Count; i++)
            {
                Assert.AreEqual(usedStyleList[i], testUsedStyleList[i]);
            }
        }

        [Test]
        public void TestAddStylesToDocument()
        {
            XWPFDocument docOut = new XWPFDocument();
            XWPFStyles styles = docOut.CreateStyles();

            String strStyleId = "headline1";
            CT_Style ctStyle = new CT_Style();

            ctStyle.styleId = (strStyleId);
            XWPFStyle s = new XWPFStyle(ctStyle);
            styles.AddStyle(s);

            Assert.IsTrue(styles.StyleExist(strStyleId));

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            styles = docIn.GetStyles();
            Assert.IsTrue(styles.StyleExist(strStyleId));
        }

        /**
         * Bug #52449 - We should be able to write a file containing
         *  both regular and glossary styles without error
         */
        [Test]
        public void Test52449()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("52449.docx");
            XWPFStyles styles = doc.GetStyles();
            Assert.IsNotNull(styles);

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            styles = docIn.GetStyles();
            Assert.IsNotNull(styles);
        }


        /**
         * YK: tests below don't make much sense,
         * they exist only to copy xml beans to pi-ooxml-schemas.jar
         */
        [Test]
        public void TestLanguages()
        {
            XWPFDocument docOut = new XWPFDocument();
            XWPFStyles styles = docOut.CreateStyles();
            styles.SetEastAsia("Chinese");

            styles.SetSpellingLanguage("English");

            CT_Fonts def = new CT_Fonts();
            styles.SetDefaultFonts(def);
        }
        [Test]
        public void TestType()
        {
            CT_Style ctStyle = new CT_Style();
            XWPFStyle style = new XWPFStyle(ctStyle);

            style.StyleType = ST_StyleType.paragraph;
            Assert.AreEqual(ST_StyleType.paragraph, style.StyleType);
        }
        [Test]
        public void TestLatentStyles()
        {
            CT_LatentStyles latentStyles = new CT_LatentStyles();
            CT_LsdException ex = latentStyles.AddNewLsdException();
            ex.name=("ex1");
            XWPFLatentStyles ls = new XWPFLatentStyles(latentStyles);
            Assert.AreEqual(true, ls.IsLatentStyle("ex1"));
            Assert.AreEqual(false, ls.IsLatentStyle("notex1"));
        }

        [Test]
        public void TestSetStyles_Bug57254()
        {
            XWPFDocument docOut = new XWPFDocument();
            XWPFStyles styles = docOut.CreateStyles();

            CT_Styles ctStyles = new CT_Styles();
            String strStyleId = "headline1";
            CT_Style ctStyle = ctStyles.AddNewStyle();

            ctStyle.styleId = (/*setter*/strStyleId);
            styles.SetStyles(ctStyles);

            Assert.IsTrue(styles.StyleExist(strStyleId));

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            styles = docIn.GetStyles();
            Assert.IsTrue(styles.StyleExist(strStyleId));
        }

        [Test]
        public void TestEasyAccessToStyles()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
            XWPFStyles styles = doc.GetStyles();
            Assert.IsNotNull(styles);

            // Has 3 paragraphs on page one, a break, and 3 on page 2
            Assert.AreEqual(7, doc.Paragraphs.Count);

            // Check the first three have no run styles, just default paragraph style
            for (int i = 0; i < 3; i++)
            {
                XWPFParagraph p = doc.Paragraphs[(i)];
                Assert.AreEqual(null, p.Style);
                Assert.AreEqual(null, p.StyleID);
                Assert.AreEqual(1, p.Runs.Count);

                XWPFRun r = p.Runs[(0)];
                Assert.AreEqual(null, r.GetColor());
                Assert.AreEqual(null, r.FontFamily);
                Assert.AreEqual(null, r.FontName);
                Assert.AreEqual(-1, r.FontSize);
            }

            // On page two, has explicit styles, but on Runs not on
            //  the paragraph itself
            for (int i = 4; i < 7; i++)
            {
                XWPFParagraph p = doc.Paragraphs[(i)];
                Assert.AreEqual(null, p.Style);
                Assert.AreEqual(null, p.StyleID);
                Assert.AreEqual(1, p.Runs.Count);

                XWPFRun r = p.Runs[(0)];
                Assert.AreEqual("Arial Black", r.FontFamily);
                Assert.AreEqual("Arial Black", r.FontName);
                Assert.AreEqual(16, r.FontSize);
                Assert.AreEqual("548DD4", r.GetColor());
            }

            // Check the document styles
            // Should have a style defined for each type
            Assert.AreEqual(4, styles.NumberOfStyles);
            Assert.IsNotNull(styles.GetStyle("Normal"));
            Assert.IsNotNull(styles.GetStyle("DefaultParagraphFont"));
            Assert.IsNotNull(styles.GetStyle("TableNormal"));
            Assert.IsNotNull(styles.GetStyle("NoList"));

            // We can't do much yet with latent styles
            Assert.AreEqual(137, styles.LatentStyles.NumberOfStyles);

            // Check the default styles
            Assert.IsNotNull(styles.DefaultRunStyle);
            Assert.IsNotNull(styles.DefaultParagraphStyle);

            Assert.AreEqual(11, styles.DefaultRunStyle.FontSize);
            Assert.AreEqual(200, styles.DefaultParagraphStyle.SpacingAfter);
        }

        // Bug 60329: style with missing StyleID throws NPE
        [Test]
        public void TestMissingStyleId()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("60329.docx");
            XWPFStyles styles = doc.GetStyles();
            // Styles exist in the test document in this order, EmptyCellLayoutStyle
            // is missing a StyleId
            try
            {
                Assert.IsNotNull(styles.GetStyle("NoList"));
                Assert.IsNull(styles.GetStyle("EmptyCellLayoutStyle"));
                Assert.IsNotNull(styles.GetStyle("BalloonText"));

                // Bug 64600: styleExist throws NPE
                Assert.IsTrue(styles.StyleExist("NoList"));
                Assert.IsFalse(styles.StyleExist("EmptyCellLayoutStyle"));
                Assert.IsTrue(styles.StyleExist("BalloonText"));
            }
            catch (NullReferenceException e)
            {
                Assert.Fail(e.ToString());
            }

            doc.Close();
        }
    }
}

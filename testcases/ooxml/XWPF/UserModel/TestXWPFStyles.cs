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
    using NUnit.Framework;using NUnit.Framework.Legacy;
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
            
            //ClassicAssert.AreEqual(usedStyleList, testUsedStyleList);
            ClassicAssert.AreEqual(usedStyleList.Count, testUsedStyleList.Count);
            for (int i = 0; i < usedStyleList.Count; i++)
            {
                ClassicAssert.AreEqual(usedStyleList[i], testUsedStyleList[i]);
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

            ClassicAssert.IsTrue(styles.StyleExist(strStyleId));

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            styles = docIn.GetStyles();
            ClassicAssert.IsTrue(styles.StyleExist(strStyleId));
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
            ClassicAssert.IsNotNull(styles);

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            styles = docIn.GetStyles();
            ClassicAssert.IsNotNull(styles);
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
            ClassicAssert.AreEqual(ST_StyleType.paragraph, style.StyleType);
        }
        [Test]
        public void TestLatentStyles()
        {
            CT_LatentStyles latentStyles = new CT_LatentStyles();
            CT_LsdException ex = latentStyles.AddNewLsdException();
            ex.name=("ex1");
            XWPFLatentStyles ls = new XWPFLatentStyles(latentStyles);
            ClassicAssert.AreEqual(true, ls.IsLatentStyle("ex1"));
            ClassicAssert.AreEqual(false, ls.IsLatentStyle("notex1"));
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

            ClassicAssert.IsTrue(styles.StyleExist(strStyleId));

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            styles = docIn.GetStyles();
            ClassicAssert.IsTrue(styles.StyleExist(strStyleId));
        }

        [Test]
        public void TestEasyAccessToStyles()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
            XWPFStyles styles = doc.GetStyles();
            ClassicAssert.IsNotNull(styles);

            // Has 3 paragraphs on page one, a break, and 3 on page 2
            ClassicAssert.AreEqual(7, doc.Paragraphs.Count);

            // Check the first three have no run styles, just default paragraph style
            for (int i = 0; i < 3; i++)
            {
                XWPFParagraph p = doc.Paragraphs[(i)];
                ClassicAssert.AreEqual(null, p.Style);
                ClassicAssert.AreEqual(null, p.StyleID);
                ClassicAssert.AreEqual(1, p.Runs.Count);

                XWPFRun r = p.Runs[(0)];
                ClassicAssert.AreEqual(null, r.GetColor());
                ClassicAssert.AreEqual(null, r.FontFamily);
                ClassicAssert.AreEqual(null, r.FontName);
                ClassicAssert.AreEqual(-1, r.FontSize);
            }

            // On page two, has explicit styles, but on Runs not on
            //  the paragraph itself
            for (int i = 4; i < 7; i++)
            {
                XWPFParagraph p = doc.Paragraphs[(i)];
                ClassicAssert.AreEqual(null, p.Style);
                ClassicAssert.AreEqual(null, p.StyleID);
                ClassicAssert.AreEqual(1, p.Runs.Count);

                XWPFRun r = p.Runs[(0)];
                ClassicAssert.AreEqual("Arial Black", r.FontFamily);
                ClassicAssert.AreEqual("Arial Black", r.FontName);
                ClassicAssert.AreEqual(16, r.FontSize);
                ClassicAssert.AreEqual("548DD4", r.GetColor());
            }

            // Check the document styles
            // Should have a style defined for each type
            ClassicAssert.AreEqual(4, styles.NumberOfStyles);
            ClassicAssert.IsNotNull(styles.GetStyle("Normal"));
            ClassicAssert.IsNotNull(styles.GetStyle("DefaultParagraphFont"));
            ClassicAssert.IsNotNull(styles.GetStyle("TableNormal"));
            ClassicAssert.IsNotNull(styles.GetStyle("NoList"));

            // We can't do much yet with latent styles
            ClassicAssert.AreEqual(137, styles.LatentStyles.NumberOfStyles);

            // Check the default styles
            ClassicAssert.IsNotNull(styles.DefaultRunStyle);
            ClassicAssert.IsNotNull(styles.DefaultParagraphStyle);

            ClassicAssert.AreEqual(11, styles.DefaultRunStyle.FontSize);
            ClassicAssert.AreEqual(200, styles.DefaultParagraphStyle.SpacingAfter);
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
                ClassicAssert.IsNotNull(styles.GetStyle("NoList"));
                ClassicAssert.IsNull(styles.GetStyle("EmptyCellLayoutStyle"));
                ClassicAssert.IsNotNull(styles.GetStyle("BalloonText"));

                // Bug 64600: styleExist throws NPE
                ClassicAssert.IsTrue(styles.StyleExist("NoList"));
                ClassicAssert.IsFalse(styles.StyleExist("EmptyCellLayoutStyle"));
                ClassicAssert.IsTrue(styles.StyleExist("BalloonText"));
            }
            catch (NullReferenceException e)
            {
                Assert.Fail(e.ToString());
            }

            doc.Close();
        }
    }
}

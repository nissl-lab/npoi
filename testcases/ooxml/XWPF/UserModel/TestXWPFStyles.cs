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

            String strStyleName = "headline1";
            CT_Style ctStyle = new CT_Style();

            ctStyle.styleId = (strStyleName);
            XWPFStyle s = new XWPFStyle(ctStyle);
            styles.AddStyle(s);

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            styles = docIn.GetStyles();
            Assert.IsTrue(styles.StyleExist(strStyleName));
        }

        //	protected void tearDown()  {
        //		super.TearDown();
        //	}

    }

}
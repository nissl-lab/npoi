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
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    [TestFixture]
    public class TestXWPFNumbering
    {

        [Test]
        public void TestCompareAbstractNum()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");
            XWPFNumbering numbering = doc.GetNumbering();
            int numId = 1;
            ClassicAssert.IsTrue(numbering.NumExist(numId.ToString()));
            XWPFNum num = numbering.GetNum(numId.ToString());
            string abstrNumId = num.GetCTNum().abstractNumId.val;
            XWPFAbstractNum abstractNum = numbering.GetAbstractNum(abstrNumId);
            string CompareAbstractNum = numbering.GetIdOfAbstractNum(abstractNum);
            ClassicAssert.AreEqual(abstrNumId, CompareAbstractNum);
        }

        [Test]
        public void TestAddNumberingToDoc()
        {
            string abstractNumId = "1";
            string numId = "1";

            XWPFDocument docOut = new XWPFDocument();
            XWPFNumbering numbering = docOut.CreateNumbering();
            numId = numbering.AddNum(abstractNumId);

            XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

            numbering = docIn.GetNumbering();
            ClassicAssert.IsTrue(numbering.NumExist(numId));
            XWPFNum num = numbering.GetNum(numId);

            string CompareAbstractNum = num.GetCTNum().abstractNumId.val;
            ClassicAssert.AreEqual(abstractNumId, CompareAbstractNum);
        }
        [Test]
        public void TestGetNumIlvl()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");
            string numIlvl = "0";
            ClassicAssert.AreEqual(numIlvl, doc.Paragraphs[0].GetNumIlvl());
            numIlvl = "1";
            ClassicAssert.AreEqual(numIlvl, doc.Paragraphs[5].GetNumIlvl());
        }
        [Test]
        public void TestGetNumFmt()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");
            ClassicAssert.AreEqual("bullet", doc.Paragraphs[0].GetNumFmt());
            ClassicAssert.AreEqual("bullet", doc.Paragraphs[1].GetNumFmt());
            ClassicAssert.AreEqual("bullet", doc.Paragraphs[2].GetNumFmt());
            ClassicAssert.AreEqual("bullet", doc.Paragraphs[3].GetNumFmt());
            ClassicAssert.AreEqual("decimal", doc.Paragraphs[4].GetNumFmt());
            ClassicAssert.AreEqual("lowerLetter", doc.Paragraphs[5].GetNumFmt());
            ClassicAssert.AreEqual("lowerRoman", doc.Paragraphs[6].GetNumFmt());
        }

        [Test]
        public void TestLvlText()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");

            ClassicAssert.AreEqual("%1.%2.%3.", doc.Paragraphs[(12)].NumLevelText);

            ClassicAssert.AreEqual("NEW-%1-FORMAT", doc.Paragraphs[(14)].NumLevelText);

            XWPFParagraph p = doc.Paragraphs[(18)];
            ClassicAssert.AreEqual("%1.", p.NumLevelText);
            //test that null doesn't throw NPE
            //ClassicAssert.IsNull(p.GetNumFmt());

            //C# enum is never null
            ClassicAssert.AreEqual(ST_NumberFormat.@decimal.ToString(), p.GetNumFmt());
        }

        [Test]
        public void TestOverrideList()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("NumberingWOverrides.docx");
            XWPFParagraph p = doc.Paragraphs[(4)];
            XWPFNumbering numbering = doc.GetNumbering();
            CT_Num ctNum = numbering.GetNum(p.GetNumID()).GetCTNum();
            ClassicAssert.AreEqual(9, ctNum.SizeOfLvlOverrideArray());
            CT_NumLvl ctNumLvl = ctNum.GetLvlOverrideArray(0);
            ClassicAssert.AreEqual("upperLetter", ctNumLvl.lvl.numFmt.val.ToString());
        }

    }
}

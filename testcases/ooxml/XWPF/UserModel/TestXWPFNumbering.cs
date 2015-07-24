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
    using NPOI.OpenXmlFormats.Wordprocessing;

    [TestFixture]
    public class TestXWPFNumbering
    {

        [Test]
        public void TestCompareAbstractNum()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");
            XWPFNumbering numbering = doc.GetNumbering();
            int numId = 1;
            Assert.IsTrue(numbering.NumExist(numId.ToString()));
            XWPFNum num = numbering.GetNum(numId.ToString());
            string abstrNumId = num.GetCTNum().abstractNumId.val;
            XWPFAbstractNum abstractNum = numbering.GetAbstractNum(abstrNumId);
            string CompareAbstractNum = numbering.GetIdOfAbstractNum(abstractNum);
            Assert.AreEqual(abstrNumId, CompareAbstractNum);
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
            Assert.IsTrue(numbering.NumExist(numId));
            XWPFNum num = numbering.GetNum(numId);

            string CompareAbstractNum = num.GetCTNum().abstractNumId.val;
            Assert.AreEqual(abstractNumId, CompareAbstractNum);
        }
        [Test]
        public void TestGetNumIlvl()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");
            string numIlvl = "0";
            Assert.AreEqual(numIlvl, doc.Paragraphs[0].GetNumIlvl());
            numIlvl = "1";
            Assert.AreEqual(numIlvl, doc.Paragraphs[5].GetNumIlvl());
        }
        [Test]
        public void TestGetNumFmt()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");
            Assert.AreEqual("bullet", doc.Paragraphs[0].GetNumFmt());
            Assert.AreEqual("bullet", doc.Paragraphs[1].GetNumFmt());
            Assert.AreEqual("bullet", doc.Paragraphs[2].GetNumFmt());
            Assert.AreEqual("bullet", doc.Paragraphs[3].GetNumFmt());
            Assert.AreEqual("decimal", doc.Paragraphs[4].GetNumFmt());
            Assert.AreEqual("lowerLetter", doc.Paragraphs[5].GetNumFmt());
            Assert.AreEqual("lowerRoman", doc.Paragraphs[6].GetNumFmt());
        }

        [Test]
        public void TestLvlText()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");

            Assert.AreEqual("%1.%2.%3.", doc.Paragraphs[(12)].NumLevelText);

            Assert.AreEqual("NEW-%1-FORMAT", doc.Paragraphs[(14)].NumLevelText);

            XWPFParagraph p = doc.Paragraphs[(18)];
            Assert.AreEqual("%1.", p.NumLevelText);
            //test that null doesn't throw NPE
            //Assert.IsNull(p.GetNumFmt());

            //C# enum is never null
            Assert.AreEqual(ST_NumberFormat.@decimal.ToString(), p.GetNumFmt());
        }

        [Test]
        public void TestOverrideList()
        {
            //TODO: for now the try/catch block ensures loading/inclusion of CTNumLevel
            //for down stream Processing.
            //Ideally, we should find files that actually use overrides and test against those.
            //Use XWPFParagraph's GetNumStartOverride() in the actual tests

            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Numbering.docx");
            XWPFParagraph p = doc.Paragraphs[(18)]; XWPFNumbering numbering = doc.CreateNumbering();
            bool ex = false;
            Assert.IsNull(p.GetNumStartOverride());
            try
            {
                numbering.GetNum(p.GetNumID()).GetCTNum().GetLvlOverrideArray(1);
            }
            catch (ArgumentOutOfRangeException e)
            {
                ex = true;
            }
            Assert.IsTrue(ex);
        }

    }
}

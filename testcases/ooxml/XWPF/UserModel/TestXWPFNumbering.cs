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

    }
}


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


namespace TestCases
{
    using NPOI;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;

    /**
     * Class to Test that we handle embeded bits in
     *  OOXML files properly
     */
    [TestFixture]
    public class TestEmbeded
    {
        [Test]
        public void TestExcel()
        {
            POIXMLDocument doc = new XSSFWorkbook(
                    POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("ExcelWithAttachments.xlsm")
            );
            Test(doc, 4);
        }
        [Test]
        public void TestWord()
        {
            POIXMLDocument doc = new XWPFDocument(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("WordWithAttachments.docx")
            );
            Test(doc, 5);
        }
        /*
        [Test]
        public void TestPowerPoint()
        {
            POIXMLDocument doc = new XSLFSlideShow(OPCPackage.Open(
                    POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("PPTWithAttachments.pptm"))
            );
            Test(doc, 4);
        }*/

        private void Test(POIXMLDocument doc, int expectedCount)
        {
            Assert.IsNotNull(doc.GetAllEmbedds());
            Assert.AreEqual(expectedCount, doc.GetAllEmbedds().Count);

            for (int i = 0; i < doc.GetAllEmbedds().Count; i++)
            {
                PackagePart pp = doc.GetAllEmbedds()[i];
                Assert.IsNotNull(pp);

                byte[] b = IOUtils.ToByteArray(pp.GetStream(System.IO.FileMode.Open));
                Assert.IsTrue(b.Length > 0);
            }
        }
    }
}





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


namespace NPOI
{

    using NPOI.UTIL.IOUtils;
    using NPOI.XSLF.XSLFSlideShow;
    using NPOI.XSSF.Usermodel.XSSFWorkbook;
    using NPOI.XWPF.usermodel.XWPFDocument;

    /**
     * Class to Test that we handle embeded bits in
     *  OOXML files properly
     */
    public class TestEmbeded
    {
        public void TestExcel()
        {
            POIXMLDocument doc = new XSSFWorkbook(
                    POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("ExcelWithAttachments.xlsm")
            );
            test(doc, 4);
        }

        public void TestWord()
        {
            POIXMLDocument doc = new XWPFDocument(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("WordWithAttachments.docx")
            );
            test(doc, 5);
        }

        public void TestPowerPoint()
        {
            POIXMLDocument doc = new XSLFSlideShow(OPCPackage.Open(
                    POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("PPTWithAttachments.pptm"))
            );
            test(doc, 4);
        }

        private void Test(POIXMLDocument doc, int expectedCount)
        {
            assertNotNull(doc.GetAllEmbedds());
            Assert.AreEqual(expectedCount, doc.GetAllEmbedds().Count);

            for (int i = 0; i < doc.GetAllEmbedds().Count; i++)
            {
                PackagePart pp = doc.GetAllEmbedds().Get(i);
                assertNotNull(pp);

                byte[] b = IOUtils.ToArray(pp.GetStream());
                Assert.IsTrue(b.Length > 0);
            }
        }
    }
}




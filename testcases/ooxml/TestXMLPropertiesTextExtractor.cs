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

    using junit.framework.TestCase;

    using org.apache.poi.Openxml4j.opc.OPCPackage;
    using NPOI.UTIL.PackageHelper;
    using NPOI.XSLF.XSLFSlideShow;
    using NPOI.XSSF.extractor.XSSFExcelExtractor;
    using NPOI.XSSF.usermodel.XSSFWorkbook;

    public class TestXMLPropertiesTextExtractor
    {
        private static POIDataSamples _ssSamples = POIDataSamples.GetSpreadSheetInstance();
        private static POIDataSamples _slSamples = POIDataSamples.GetSlideShowInstance();
        [TestMethod]
        public void TestGetFromMainExtractor()
        {
            OPCPackage pkg = PackageHelper.Open(_ssSamples.OpenResourceAsStream("ExcelWithAttachments.xlsm"));

            XSSFWorkbook wb = new XSSFWorkbook(pkg);

            XSSFExcelExtractor ext = new XSSFExcelExtractor(wb);
            POIXMLPropertiesTextExtractor textExt = ext.GetMetadataTextExtractor();

            // Check basics
            assertNotNull(textExt);
            Assert.IsTrue(textExt.GetText().Length > 0);

            // Check some of the content
            String text = textExt.GetText();
            String cText = textExt.GetCorePropertiesText();

            Assert.IsTrue(text.Contains("LastModifiedBy = Yury Batrakov"));
            Assert.IsTrue(cText.Contains("LastModifiedBy = Yury Batrakov"));

            textExt.Close();
            ext.Close();
        }
        [TestMethod]
        public void TestCore()
        {
            OPCPackage pkg = PackageHelper.Open(
                    _ssSamples.OpenResourceAsStream("ExcelWithAttachments.xlsm")
            );
            XSSFWorkbook wb = new XSSFWorkbook(pkg);

            POIXMLPropertiesTextExtractor ext = new POIXMLPropertiesTextExtractor(wb);
            ext.GetText();

            // Now check
            String text = ext.GetText();
            String cText = ext.GetCorePropertiesText();

            Assert.IsTrue(text.Contains("LastModifiedBy = Yury Batrakov"));
            Assert.IsTrue(cText.Contains("LastModifiedBy = Yury Batrakov"));

            ext.Close();
        }
        [TestMethod]
        public void TestExtended()
        {
            OPCPackage pkg = OPCPackage.Open(
                    _ssSamples.OpenResourceAsStream("ExcelWithAttachments.xlsm")
            );
            XSSFWorkbook wb = new XSSFWorkbook(pkg);

            POIXMLPropertiesTextExtractor ext = new POIXMLPropertiesTextExtractor(wb);
            ext.GetText();

            // Now check
            String text = ext.GetText();
            String eText = ext.GetExtendedPropertiesText();

            Assert.IsTrue(text.Contains("Application = Microsoft Excel"));
            Assert.IsTrue(text.Contains("Company = Mera"));
            Assert.IsTrue(eText.Contains("Application = Microsoft Excel"));
            Assert.IsTrue(eText.Contains("Company = Mera"));

            ext.Close();
        }

        public void TestCustom()
        {
            // TODO!
        }

        /**
         * Bug #49386 - some properties, especially
         *  dates can be null
         */
        [TestMethod]
        public void TestWithSomeNulls()
        {
            OPCPackage pkg = OPCPackage.Open(
                  _slSamples.OpenResourceAsStream("49386-null_dates.pptx")
            );
            XSLFSlideShow sl = new XSLFSlideShow(pkg);

            POIXMLPropertiesTextExtractor ext = new POIXMLPropertiesTextExtractor(sl);
            ext.GetText();

            String text = ext.GetText();
            Assert.IsFalse(text.Contains("Created =")); // With date is null
            Assert.IsTrue(text.Contains("CreatedString = ")); // Via string is blank
            Assert.IsTrue(text.Contains("LastModifiedBy = IT Client Services"));

            ext.Close();
        }
    }

}


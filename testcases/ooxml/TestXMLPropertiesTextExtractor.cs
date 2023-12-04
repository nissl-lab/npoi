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
    using System;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using NPOI.XSSF.Extractor;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using TestCases;

    [TestFixture]
    public class TestXMLPropertiesTextExtractor
    {
        private static POIDataSamples _ssSamples = POIDataSamples.GetSpreadSheetInstance();
        private static POIDataSamples _slSamples = POIDataSamples.GetSlideShowInstance();

        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestGetFromMainExtractor()
        {
            OPCPackage pkg = PackageHelper.Open(_ssSamples.OpenResourceAsStream("ExcelWithAttachments.xlsm"));

            XSSFWorkbook wb = new XSSFWorkbook(pkg);

            XSSFExcelExtractor ext = new XSSFExcelExtractor(wb);
            POIXMLPropertiesTextExtractor textExt = (POIXMLPropertiesTextExtractor)ext.MetadataTextExtractor;

            // Check basics
            Assert.IsNotNull(textExt);
            Assert.IsTrue(textExt.Text.Length > 0);

            // Check some of the content
            String text = textExt.Text;
            String cText = textExt.GetCorePropertiesText();

            Assert.IsTrue(text.Contains("LastModifiedBy = Yury Batrakov"));
            Assert.IsTrue(cText.Contains("LastModifiedBy = Yury Batrakov"));

            textExt.Close();
            ext.Close();
        }
        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestCore()
        {
            OPCPackage pkg = PackageHelper.Open(
                    _ssSamples.OpenResourceAsStream("ExcelWithAttachments.xlsm")
            );
            XSSFWorkbook wb = new XSSFWorkbook(pkg);

            POIXMLPropertiesTextExtractor ext = new POIXMLPropertiesTextExtractor(wb);

            // Now check
            String text = ext.Text;
            String cText = ext.GetCorePropertiesText();

            Assert.IsTrue(text.Contains("LastModifiedBy = Yury Batrakov"));
            Assert.IsTrue(cText.Contains("LastModifiedBy = Yury Batrakov"));

            ext.Close();
        }
        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestExtended()
        {
            OPCPackage pkg = OPCPackage.Open(
                    _ssSamples.OpenResourceAsStream("ExcelWithAttachments.xlsm")
            );
            XSSFWorkbook wb = new XSSFWorkbook(pkg);

            POIXMLPropertiesTextExtractor ext = new POIXMLPropertiesTextExtractor(wb);

            // Now check
            String text = ext.Text;
            String eText = ext.GetExtendedPropertiesText();

            Assert.IsTrue(text.Contains("Application = Microsoft Excel"));
            Assert.IsTrue(text.Contains("Company = Mera"));
            Assert.IsTrue(eText.Contains("Application = Microsoft Excel"));
            Assert.IsTrue(eText.Contains("Company = Mera"));

            ext.Close();
        }

    }

}


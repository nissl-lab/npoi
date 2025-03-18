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

namespace TestCases.HPSF.Extractor
{
    using NPOI.HPSF;
    using NPOI.HPSF.Extractor;
    using NPOI.HSSF.Extractor;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using NUnit.Framework;
    using System;
    using System.IO;
    using TestCases.HSSF;

    [TestFixture]
    public class TestHPSFPropertiesExtractor : POITestCase
    {
        private static POIDataSamples _samples = POIDataSamples.GetHPSFInstance();

        [Test]
        public void TestNormalProperties()
        {
            POIFSFileSystem fs = new POIFSFileSystem(_samples.OpenResourceAsStream("TestMickey.doc"));
            HPSFPropertiesExtractor ext = new HPSFPropertiesExtractor(fs);
            String text = ext.Text;

            // Check each bit in turn
            String summary = ext.SummaryInformationText;
            String docsummary = ext.DocumentSummaryInformationText;

            AssertContains(summary, "TEMPLATE = Normal");
            AssertContains(summary, "SUBJECT = sample subject");
            AssertContains(docsummary, "MANAGER = sample manager");
            AssertContains(docsummary, "COMPANY = sample company");

            // Now overall
            text = ext.Text;
            AssertContains(text, "TEMPLATE = Normal");
            AssertContains(text, "SUBJECT = sample subject");
            AssertContains(text, "MANAGER = sample manager");
            AssertContains(text, "COMPANY = sample company");
        }

        [Test]
        public void TestNormalUnicodeProperties()
        {
            POIFSFileSystem fs = new POIFSFileSystem(_samples.OpenResourceAsStream("TestUnicode.xls"));
            HPSFPropertiesExtractor ext = new HPSFPropertiesExtractor(fs);

            // Check each bit in turn
            String summary = ext.SummaryInformationText;
            String docsummary = ext.DocumentSummaryInformationText;

            AssertContains(summary, "AUTHOR = marshall");
            AssertContains(summary, "TITLE = Titel: \u00c4h");
            AssertContains(docsummary, "COMPANY = Schreiner");
            AssertContains(docsummary, "SCALE = False");

            // Now overall
            string text = ext.Text;
            AssertContains(text, "AUTHOR = marshall");
            AssertContains(text, "TITLE = Titel: \u00c4h");
            AssertContains(text, "COMPANY = Schreiner");
            AssertContains(text, "SCALE = False");
        }

        [Test]
        public void TestCustomProperties()
        {
            POIFSFileSystem fs = new POIFSFileSystem(
                    _samples.OpenResourceAsStream("TestMickey.doc")
            );
            HPSFPropertiesExtractor ext = new HPSFPropertiesExtractor(fs);

            // Custom properties are part of the document info stream
            String dinfText = ext.DocumentSummaryInformationText;
            AssertContains(dinfText, "Client = sample client");
            AssertContains(dinfText, "Division = sample division");

            String text = ext.Text;
            AssertContains(text, "Client = sample client");
            AssertContains(text, "Division = sample division");
        }

        [Test]
        public void TestConstructors()
        {
            POIFSFileSystem fs;
            HSSFWorkbook wb;
            try
            {
                fs = new POIFSFileSystem(_samples.OpenResourceAsStream("TestUnicode.xls"));
                wb = new HSSFWorkbook(fs);
            }
            catch (IOException e)
            {
                throw new Exception("TestConstructors", e);
            }
            ExcelExtractor excelExt = new ExcelExtractor(wb);

            String fsText;
            HPSFPropertiesExtractor fsExt = new HPSFPropertiesExtractor(fs);
            fsExt.SetFilesystem(null); // Don't close re-used test resources!
            try
            {
                fsText = fsExt.Text;
            }
            finally
            {
                fsExt.Close();
            }

            String hwText;
            HPSFPropertiesExtractor hwExt = new HPSFPropertiesExtractor(wb);
            hwExt.SetFilesystem(null); // Don't close re-used test resources!
            try
            {
                hwText = hwExt.Text;
            }
            finally
            {
                hwExt.Close();
            }

            String eeText;
            HPSFPropertiesExtractor eeExt = new HPSFPropertiesExtractor(excelExt);
            eeExt.SetFilesystem(null); // Don't close re-used test resources!
            try
            {
                eeText = eeExt.Text;
            }
            finally
            {
                eeExt.Close();
                wb.Close();
            }

            Assert.AreEqual(fsText, hwText);
            Assert.AreEqual(fsText, eeText);

            AssertContains(fsText, "AUTHOR = marshall");
            AssertContains(fsText, "TITLE = Titel: \u00c4h");
        }

        [Test]
        public void Test42726()
        {
            HPSFPropertiesExtractor ex = new HPSFPropertiesExtractor(HSSFTestDataSamples.OpenSampleWorkbook("42726.xls"));
            String txt = ex.Text;
            AssertContains(txt, "PID_AUTHOR");
            AssertContains(txt, "PID_EDITTIME");
            AssertContains(txt, "PID_REVNUMBER");
            AssertContains(txt, "PID_THUMBNAIL");
        }

        [Test]
        public void TestThumbnail()
        {
            POIFSFileSystem fs = new POIFSFileSystem(_samples.OpenResourceAsStream("TestThumbnail.xls"));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            Thumbnail thumbnail = new Thumbnail(wb.SummaryInformation.Thumbnail);
            Assert.AreEqual(-1, thumbnail.ClipboardFormatTag);
            Assert.AreEqual(3, thumbnail.GetClipboardFormat());
            Assert.IsNotNull(thumbnail.GetThumbnailAsWMF());
            //wb.Close();
        }

        [Test]
        public void Test52258()
        {
            POIFSFileSystem fs = new POIFSFileSystem(_samples.OpenResourceAsStream("TestVisioWithCodepage.vsd"));
            HPSFPropertiesExtractor ext = new HPSFPropertiesExtractor(fs);
            try
            {
                Assert.IsNotNull(ext.DocSummaryInformation);
                Assert.IsNotNull(ext.DocumentSummaryInformationText);
                Assert.IsNotNull(ext.SummaryInformation);
                Assert.IsNotNull(ext.SummaryInformationText);
                Assert.IsNotNull(ext.Text);
            }
            finally
            {
                ext.Close();
            }
        }
    }


}
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
    using System;



    using NUnit.Framework;

    using NPOI.HSSF;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using System.IO;
    using NPOI.HSSF.Extractor;
    using NPOI.HPSF.Extractor;
    using TestCases.HSSF;
    using NPOI.HPSF;

    [TestFixture]
    public class TestHPSFPropertiesExtractor
    {
        private static POIDataSamples _samples = POIDataSamples.GetHPSFInstance();

        [Test]
        public void TestNormalProperties()
        {
            POIFSFileSystem fs = new POIFSFileSystem(_samples.OpenResourceAsStream("TestMickey.doc"));
            HPSFPropertiesExtractor ext = new HPSFPropertiesExtractor(fs);
            String text = ext.Text;

            // Check each bit in turn
            String sinfText = ext.SummaryInformationText;
            String dinfText = ext.DocumentSummaryInformationText;

            Assert.IsTrue(sinfText.IndexOf("TEMPLATE = Normal") > -1);
            Assert.IsTrue(sinfText.IndexOf("SUBJECT = sample subject") > -1);
            Assert.IsTrue(dinfText.IndexOf("MANAGER = sample manager") > -1);
            Assert.IsTrue(dinfText.IndexOf("COMPANY = sample company") > -1);

            // Now overall
             text = ext.Text;
            Assert.IsTrue(text.IndexOf("TEMPLATE = Normal") > -1);
            Assert.IsTrue(text.IndexOf("SUBJECT = sample subject") > -1);
            Assert.IsTrue(text.IndexOf("MANAGER = sample manager") > -1);
            Assert.IsTrue(text.IndexOf("COMPANY = sample company") > -1);
        }

        [Test]
        public void TestNormalUnicodeProperties()
        {
            POIFSFileSystem fs = new POIFSFileSystem(_samples.OpenResourceAsStream("TestUnicode.xls"));
            HPSFPropertiesExtractor ext = new HPSFPropertiesExtractor(fs);
            string text = ext.Text;

            // Check each bit in turn
            String sinfText = ext.SummaryInformationText;
            String dinfText = ext.DocumentSummaryInformationText;

            Assert.IsTrue(sinfText.IndexOf("AUTHOR = marshall") > -1);
            Assert.IsTrue(sinfText.IndexOf("TITLE = Titel: \u00c4h") > -1);
            Assert.IsTrue(dinfText.IndexOf("COMPANY = Schreiner") > -1);
            Assert.IsTrue(dinfText.IndexOf("SCALE = False") > -1);

            // Now overall
            text = ext.Text;
            Assert.IsTrue(text.IndexOf("AUTHOR = marshall") > -1);
            Assert.IsTrue(text.IndexOf("TITLE = Titel: \u00c4h") > -1);
            Assert.IsTrue(text.IndexOf("COMPANY = Schreiner") > -1);
            Assert.IsTrue(text.IndexOf("SCALE = False") > -1);
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
            Assert.IsTrue(dinfText.IndexOf("Client = sample client") > -1);
            Assert.IsTrue(dinfText.IndexOf("Division = sample division") > -1);

            String text = ext.Text;
            Assert.IsTrue(text.IndexOf("Client = sample client") > -1);
            Assert.IsTrue(text.IndexOf("Division = sample division") > -1);
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

            String fsText = (new HPSFPropertiesExtractor(fs)).Text;
            String hwText = (new HPSFPropertiesExtractor(wb)).Text;
            String eeText = (new HPSFPropertiesExtractor(excelExt)).Text;

            Assert.AreEqual(fsText, hwText);
            Assert.AreEqual(fsText, eeText);

            Assert.IsTrue(fsText.IndexOf("AUTHOR = marshall") > -1);
            Assert.IsTrue(fsText.IndexOf("TITLE = Titel: \u00c4h") > -1);
        }

        [Test]
        public void Test42726()
        {
            HPSFPropertiesExtractor ex = new HPSFPropertiesExtractor(HSSFTestDataSamples.OpenSampleWorkbook("42726.xls"));
            String txt = ex.Text;
            Assert.IsTrue(txt.IndexOf("PID_AUTHOR") != -1);
            Assert.IsTrue(txt.IndexOf("PID_EDITTIME") != -1);
            Assert.IsTrue(txt.IndexOf("PID_REVNUMBER") != -1);
            Assert.IsTrue(txt.IndexOf("PID_THUMBNAIL") != -1);
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

    }


}
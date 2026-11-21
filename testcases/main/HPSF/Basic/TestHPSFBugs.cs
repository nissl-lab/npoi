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

namespace TestCases.HPSF.Basic
{
    using System.IO;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.HSSF.UserModel;
    using System;
    using NPOI.SS.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI.HPSF;
    using NPOI;

    /**
     * Tests various bugs have been fixed
     */
    [TestFixture]
    public class TestHPSFBugs
    {
        private static POIDataSamples _samples = POIDataSamples.GetHPSFInstance();
        /**
         * Ensure that we can create a new HSSF Workbook,
         *  then add some properties to it, save +
         *  reload, and still access & change them.
         */
        [Test]
        public void Test48832()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // Starts empty
            ClassicAssert.IsNull(wb.DocumentSummaryInformation);
            ClassicAssert.IsNull(wb.SummaryInformation);

            // Add new properties
            wb.CreateInformationProperties();

            ClassicAssert.IsNotNull(wb.DocumentSummaryInformation);
            ClassicAssert.IsNotNull(wb.SummaryInformation);

            // Set Initial values
            wb.SummaryInformation.Author = (/*setter*/"Apache POI");
            wb.SummaryInformation.Keywords = (/*setter*/"Testing POI");
            wb.SummaryInformation.CreateDateTime = DateUtil.GetJavaDate(12345);

            wb.DocumentSummaryInformation.Company = (/*setter*/"Apache");


            // Save and reload
            MemoryStream baos = new MemoryStream();
            wb.Write(baos);
            MemoryStream bais =
               new MemoryStream(baos.ToArray());
            wb = new HSSFWorkbook(bais);


            // Ensure Changes were taken
            ClassicAssert.IsNotNull(wb.DocumentSummaryInformation);
            ClassicAssert.IsNotNull(wb.SummaryInformation);

            ClassicAssert.AreEqual("Apache POI", wb.SummaryInformation.Author);
            ClassicAssert.AreEqual("Testing POI", wb.SummaryInformation.Keywords);
            ClassicAssert.AreEqual(12345, DateUtil.GetExcelDate(wb.SummaryInformation.CreateDateTime.Value));
            ClassicAssert.AreEqual("Apache", wb.DocumentSummaryInformation.Company);


            // Set some more, save + reload
            wb.SummaryInformation.Comments = (/*setter*/"Resaved");

            baos = new MemoryStream();
            wb.Write(baos);
            bais = new MemoryStream(baos.ToArray());
            wb = new HSSFWorkbook(bais);

            // Check again
            ClassicAssert.IsNotNull(wb.DocumentSummaryInformation);
            ClassicAssert.IsNotNull(wb.SummaryInformation);

            ClassicAssert.AreEqual("Apache POI", wb.SummaryInformation.Author);
            ClassicAssert.AreEqual("Testing POI", wb.SummaryInformation.Keywords);
            ClassicAssert.AreEqual("Resaved", wb.SummaryInformation.Comments);
            ClassicAssert.AreEqual(12345, DateUtil.GetExcelDate(wb.SummaryInformation.CreateDateTime.Value));
            ClassicAssert.AreEqual("Apache", wb.DocumentSummaryInformation.Company);
        }

        /**
        * Some files seem to want the length and data to be on a 4-byte boundary,
        * and without that you'll hit an ArrayIndexOutOfBoundsException after
        * reading junk
        */
        [Test]
        public void Test54233()
        {
            DocumentInputStream dis;
            NPOIFSFileSystem fs =
                    new NPOIFSFileSystem(_samples.OpenResourceAsStream("TestNon4ByteBoundary.doc"));

            dis = fs.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME);
            SummaryInformation si = (SummaryInformation)PropertySetFactory.Create(dis);

            dis = fs.CreateDocumentInputStream(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            DocumentSummaryInformation dsi = (DocumentSummaryInformation)PropertySetFactory.Create(dis);

            // Test
            ClassicAssert.AreEqual("Microsoft Word 10.0", si.ApplicationName);
            ClassicAssert.AreEqual("", si.Title);
            ClassicAssert.AreEqual("", si.Author);
            ClassicAssert.AreEqual("Cour de Justice", dsi.Company);


            // Write out and read back, should still be valid
            POIDocument doc = new HPSFPropertiesOnlyDocument(fs);
            MemoryStream baos = new MemoryStream();
            doc.Write(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());
            doc = new HPSFPropertiesOnlyDocument(new POIFSFileSystem(bais));

            // Check properties are still there
            ClassicAssert.AreEqual("Microsoft Word 10.0", si.ApplicationName);
            ClassicAssert.AreEqual("", si.Title);
            ClassicAssert.AreEqual("", si.Author);
            ClassicAssert.AreEqual("Cour de Justice", dsi.Company);
        }

        /**
        * CodePage Strings can be zero length
        */
        [Test]
        public void Test56138()
        {
            DocumentInputStream dis;
            POIFSFileSystem fs =
                    new POIFSFileSystem(_samples.OpenResourceAsStream("TestZeroLengthCodePage.mpp"));

            dis = fs.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME);
            SummaryInformation si = (SummaryInformation)PropertySetFactory.Create(dis);

            dis = fs.CreateDocumentInputStream(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            DocumentSummaryInformation dsi = (DocumentSummaryInformation)PropertySetFactory.Create(dis);

            // Test
            ClassicAssert.AreEqual("MSProject", si.ApplicationName);
            ClassicAssert.AreEqual("project1", si.Title);
            ClassicAssert.AreEqual("Jon Iles", si.Author);

            ClassicAssert.AreEqual("", dsi.Company);
            ClassicAssert.AreEqual(2, dsi.SectionCount);
        }
    }

}
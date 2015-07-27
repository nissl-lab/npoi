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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;

    using TestCases.HSSF;
    using NPOI.HPSF;
    using NPOI.POIFS.FileSystem;

    /**
     * Old-style setting of POIFS properties doesn't work with POI 3.0.2
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestPOIFSProperties
    {

        private static String title = "Testing POIFS properties";

        public void TestFail()  
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("Simple.xls");
            POIFSFileSystem fs = new POIFSFileSystem(is1);

            HSSFWorkbook wb = new HSSFWorkbook(fs);

            //set POIFS properties after constructing HSSFWorkbook
            //(a piece of code that used to work up to POI 3.0.2)
            SummaryInformation summary1 = (SummaryInformation)PropertySetFactory.Create(fs.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME));
            summary1.Title=(title);
            //Write the modified property back to POIFS
            fs.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME).Delete();
            fs.CreateDocument(summary1.ToInputStream(), SummaryInformation.DEFAULT_STREAM_NAME);

            //save the workbook and read the property
            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);
            out1.Close();

            POIFSFileSystem fs2 = new POIFSFileSystem(new MemoryStream(out1.ToArray()));
            SummaryInformation summary2 = (SummaryInformation)PropertySetFactory.Create(fs2.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME));

            //Assert.Failing assertion
            Assert.AreEqual(title, summary2.Title);
        }

        [Test]
        public void TestOK()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("Simple.xls");
            POIFSFileSystem fs = new POIFSFileSystem(is1);

            //set POIFS properties before constructing HSSFWorkbook
            SummaryInformation summary1 = (SummaryInformation)PropertySetFactory.Create(fs.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME));
            summary1.Title = (title);

            fs.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME).Delete();
            fs.CreateDocument(summary1.ToInputStream(), SummaryInformation.DEFAULT_STREAM_NAME);

            HSSFWorkbook wb = new HSSFWorkbook(fs);

            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);
            out1.Close();

            //read the property
            POIFSFileSystem fs2 = new POIFSFileSystem(new MemoryStream(out1.ToArray()));
            SummaryInformation summary2 = (SummaryInformation)PropertySetFactory.Create(fs2.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME));
            Assert.AreEqual(title, summary2.Title);

        }
    }
}
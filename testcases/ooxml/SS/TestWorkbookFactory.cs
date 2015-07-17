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

namespace NPOI.SS
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using TestCases.HSSF;
    using System.Configuration;
    using System.IO;

    [TestFixture]
    public class TestWorkbookFactory
    {
        private String xls;
        private String xlsx;
        private String txt;

        string testdataPath;
        [SetUp]
        public void SetUp()
        {
            xls = "SampleSS.xls";
            xlsx = "SampleSS.xlsx";
            txt = "SampleSS.txt";
            testdataPath = ConfigurationManager.AppSettings["POI.testdata.path"] + "\\spreadsheet\\";
        }

        [Test]
        public void TestCreateNative()
        {
            IWorkbook wb;

            // POIFS -> hssf
            wb = WorkbookFactory.Create(
                    new POIFSFileSystem(HSSFTestDataSamples.OpenSampleFileStream(xls))
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();

            // Package -> xssf
            wb = WorkbookFactory.Create(
                    OPCPackage.Open(
                            HSSFTestDataSamples.OpenSampleFileStream(xlsx))
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
        }

        /**
         * Creates the appropriate kind of Workbook, but
         *  Checking the mime magic at the start of the
         *  InputStream, then creating what's required.
         */
        [Test]
        public void TestCreateGeneric()
        {
            IWorkbook wb;

            // InputStream -> either
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xls)
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();

            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xlsx)
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            // File -> either
            wb = WorkbookFactory.Create(
                  testdataPath + xls
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();

            wb = WorkbookFactory.Create(
                  testdataPath + xlsx
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            // TODO: close() re-writes the sample-file?! Resort to revert() for now to close file handle...
            ((XSSFWorkbook)wb).Package.Revert();

            // Invalid type -> exception
            try
            {
                Stream stream = HSSFTestDataSamples.OpenSampleFileStream(txt);
                try
                {
                    wb = WorkbookFactory.Create(stream);
                }
                finally
                {
                    stream.Close();
                }
                Assert.Fail();
            }
            catch (ArgumentException e)
            {
                // Good
            }
        }
    }

}
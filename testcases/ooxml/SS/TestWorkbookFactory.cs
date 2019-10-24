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
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.Util;

    [TestFixture]
    public class TestWorkbookFactory
    {
        private String xls;
        private String xlsx;
        private String[] xls_prot;
        private String[] xlsx_prot;
        private String txt;

        string testdataPath;
        [SetUp]
        public void SetUp()
        {
            xls = "SampleSS.xls";
            xlsx = "SampleSS.xlsx";
            xls_prot = new String[] { "password.xls", "password" };
            xlsx_prot = new String[] { "protected_passtika.xlsx", "tika" };
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
        [Test]
        public void TestCreateReadOnly()
        {
            IWorkbook wb;

            // POIFS -> hssf
            wb = WorkbookFactory.Create(HSSFTestDataSamples.GetSampleFile(xls).FullName, null, true);
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();

            // Package -> xssf
            wb = WorkbookFactory.Create(HSSFTestDataSamples.GetSampleFile(xlsx).FullName, null, true);
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            wb.Close();
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
                  Path.GetFullPath(testdataPath + xls)
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
            catch (InvalidFormatException e)
            {
                // Good
            }
        }

        /**
         * Check that the overloaded stream methods which take passwords work properly
         */
        [Test]
        public void TestCreateWithPasswordFromStream()
        {
            IWorkbook wb;
            // Unprotected, no password given, opens normally
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xls), null
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xlsx), null
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            // Unprotected, wrong password, opens normally
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xls), "wrong"
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xlsx), "wrong"
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            // Protected, correct password, opens fine
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xls_prot[0]), xls_prot[1]
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xlsx_prot[0]), xlsx_prot[1]
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            // Protected, wrong password, throws Exception
            try
            {
                wb = WorkbookFactory.Create(
                        HSSFTestDataSamples.OpenSampleFileStream(xls_prot[0]), "wrong"
                );
                Assert.Fail("Shouldn't be able to open with the wrong password");
            }
            catch (EncryptedDocumentException e) { }
            try
            {
                wb = WorkbookFactory.Create(
                        HSSFTestDataSamples.OpenSampleFileStream(xlsx_prot[0]), "wrong"
                );
                Assert.Fail("Shouldn't be able to open with the wrong password");
            }
            catch (EncryptedDocumentException e) { }
        }
        /**
         * Check that the overloaded file methods which take passwords work properly
         */
        [Test]
        public void TestCreateWithPasswordFromFile()
        {
            IWorkbook wb;
            // Unprotected, no password given, opens normally
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xls).FullName, null
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xlsx).FullName, null
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            wb.Close();
            // Unprotected, wrong password, opens normally
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xls).FullName, "wrong"
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xlsx).FullName, "wrong"
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            wb.Close();
            // Protected, correct password, opens fine
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xls_prot[0]).FullName, xls_prot[1]
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            wb.Close();
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xlsx_prot[0]).FullName, xlsx_prot[1]
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            wb.Close();
            // Protected, wrong password, throws Exception
            try
            {
                wb = WorkbookFactory.Create(
                        HSSFTestDataSamples.GetSampleFile(xls_prot[0]).FullName, "wrong"
                );
                Assert.Fail("Shouldn't be able to open with the wrong password");
            }
            catch (EncryptedDocumentException e) { }
            try
            {
                wb = WorkbookFactory.Create(
                        HSSFTestDataSamples.GetSampleFile(xlsx_prot[0]).FullName, "wrong"
                );
                Assert.Fail("Shouldn't be able to open with the wrong password");
            }
            catch (EncryptedDocumentException e) { }
        }

        /**
         * Check that a helpful exception is given on an empty file / stream
         */
        [Test]
        public void TestEmptyFile()
        {
            InputStream emptyStream = new ByteArrayInputStream(new byte[0]);
            FileInfo emptyFile = TempFile.CreateTempFile("empty", ".poi");

            try
            {
                WorkbookFactory.Create(emptyStream);
                Assert.Fail("Shouldn't be able to create for an empty stream");
            }
            catch (EmptyFileException e)
            {
            }

            try
            {
                WorkbookFactory.Create(emptyFile.FullName);
                Assert.Fail("Shouldn't be able to create for an empty file");
            }
            catch (EmptyFileException e)
            {
            }
            emptyFile.Delete();
        }

    }

}
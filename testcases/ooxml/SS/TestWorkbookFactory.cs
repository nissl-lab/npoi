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

namespace TestCases.SS
{
    using NPOI;
    using NPOI.HSSF.UserModel;
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using System;
    using System.IO;
    using TestCases;
    using TestCases.HSSF;

    [TestFixture]
    public class TestWorkbookFactory
    {
        private readonly String xls = "SampleSS.xls";
        private readonly String xlsx = "SampleSS.xlsx";
        private readonly String[] xls_prot = new String[] { "password.xls", "password" };
        private readonly String[] xlsx_prot = new String[] { "protected_passtika.xlsx", "tika" };
        private readonly String txt = "SampleSS.txt";

        private readonly string testdataPath = Path.Combine(TestContext.CurrentContext.TestDirectory,
            TestContext.Parameters[POIDataSamples.TEST_PROPERTY], "spreadsheet");
        private static POILogger LOGGER = POILogFactory.GetLogger(typeof(TestWorkbookFactory));

        /**
         * Closes the sample workbook read in from filename.
         * Throws an exception if closing the workbook results in the file on disk getting modified.
         *
         * @param filename the sample workbook to read in
         * @param wb the workbook to close
         * @throws IOException
         */
        private static void AssertCloseDoesNotModifyFile(String filename, IWorkbook wb) {
            byte[] before = HSSFTestDataSamples.GetTestDataFileContent(filename);
            CloseOrRevert(wb);
            byte[] after = HSSFTestDataSamples.GetTestDataFileContent(filename);
            CollectionAssert.AreEqual(before, after, filename + " sample file was modified as a result of closing the workbook");
        }
        /**
         * // TODO: close() re-writes the sample-file?! Resort to revert() for now to close file handle...
         * Revert the changes that were made to the workbook rather than closing the workbook.
         * This allows the file handle to be closed to avoid the file handle leak detector.
         * This is a temporary fix until we figure out why wb.close() writes changes to disk.
         *
         * @param wb
         */
        private static void CloseOrRevert(IWorkbook wb)
        {
            // TODO: close() re-writes the sample-file?! Resort to revert() for now to close file handle...
            if (wb is HSSFWorkbook)
            {
                wb.Close();
            }
            else if (wb is XSSFWorkbook)
            {
                XSSFWorkbook xwb = (XSSFWorkbook)wb;
                if (PackageAccess.READ == xwb.Package.GetPackageAccess())
                {
                    xwb.Close();
                }
                else
                {
                    LOGGER.Log(POILogger.WARN,
                            "reverting XSSFWorkbook rather than closing it to avoid close() modifying the file on disk. " +
                            "Refer to bug 58779.");
                    xwb.Package.Revert();
                }
            }
            else
            {
                throw new RuntimeException("Unsupported workbook type");
            }
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
            AssertCloseDoesNotModifyFile(xls, wb);

            // Package -> xssf
            wb = WorkbookFactory.Create(
                    OPCPackage.Open(
                            HSSFTestDataSamples.OpenSampleFileStream(xlsx))
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            AssertCloseDoesNotModifyFile(xlsx, wb);
        }
        [Test]
        public void TestCreateReadOnly()
        {
            IWorkbook wb;

            // POIFS -> hssf
            wb = WorkbookFactory.Create(HSSFTestDataSamples.GetSampleFile(xls).FullName, null, true);
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            AssertCloseDoesNotModifyFile(xls, wb);

            // Package -> xssf
            wb = WorkbookFactory.Create(HSSFTestDataSamples.GetSampleFile(xlsx).FullName, null, true);
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            AssertCloseDoesNotModifyFile(xlsx, wb);
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
            AssertCloseDoesNotModifyFile(xls, wb);

            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.OpenSampleFileStream(xlsx)
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            AssertCloseDoesNotModifyFile(xlsx, wb);

            // File -> either
            wb = WorkbookFactory.Create(
                  Path.GetFullPath(Path.Combine(testdataPath, xls))
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            AssertCloseDoesNotModifyFile(xls, wb);

            wb = WorkbookFactory.Create(
                  Path.Combine(testdataPath, xlsx)
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            AssertCloseDoesNotModifyFile(xlsx, wb);

            // Invalid type -> exception
            byte[] before = HSSFTestDataSamples.GetTestDataFileContent(txt);
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
            catch (InvalidFormatException)
            {
                // Good
            }
            byte[] after = HSSFTestDataSamples.GetTestDataFileContent(txt);
            CollectionAssert.AreEqual(before, after, "Invalid type file was modified after trying to open the file as a spreadsheet");
        }

        /**
         * Check that the overloaded stream methods which take passwords work properly
         */
        //[Test]
        //public void TestCreateWithPasswordFromStream()
        //{
        //    IWorkbook wb;
        //    // Unprotected, no password given, opens normally
        //    wb = WorkbookFactory.Create(
        //            HSSFTestDataSamples.OpenSampleFileStream(xls), null
        //    );
        //    Assert.IsNotNull(wb);
        //    Assert.IsTrue(wb is HSSFWorkbook);
        //    AssertCloseDoesNotModifyFile(xls, wb);

        //    wb = WorkbookFactory.Create(
        //            HSSFTestDataSamples.OpenSampleFileStream(xlsx), null
        //    );
        //    Assert.IsNotNull(wb);
        //    Assert.IsTrue(wb is XSSFWorkbook);
        //    AssertCloseDoesNotModifyFile(xlsx, wb);

        //    // Unprotected, wrong password, opens normally
        //    wb = WorkbookFactory.Create(
        //            HSSFTestDataSamples.OpenSampleFileStream(xls), "wrong"
        //    );
        //    Assert.IsNotNull(wb);
        //    Assert.IsTrue(wb is HSSFWorkbook);
        //    AssertCloseDoesNotModifyFile(xls, wb);


        //    wb = WorkbookFactory.Create(
        //            HSSFTestDataSamples.OpenSampleFileStream(xlsx), "wrong"
        //    );
        //    Assert.IsNotNull(wb);
        //    Assert.IsTrue(wb is XSSFWorkbook);
        //    AssertCloseDoesNotModifyFile(xlsx, wb);

        //    // Protected, correct password, opens fine
        //    wb = WorkbookFactory.Create(
        //            HSSFTestDataSamples.OpenSampleFileStream(xls_prot[0]), xls_prot[1]
        //    );
        //    Assert.IsNotNull(wb);
        //    Assert.IsTrue(wb is HSSFWorkbook);
        //    AssertCloseDoesNotModifyFile(xls_prot[0], wb);

        //    wb = WorkbookFactory.Create(
        //            HSSFTestDataSamples.OpenSampleFileStream(xlsx_prot[0]), xlsx_prot[1]
        //    );
        //    Assert.IsNotNull(wb);
        //    Assert.IsTrue(wb is XSSFWorkbook);
        //    AssertCloseDoesNotModifyFile(xlsx_prot[0], wb);

        //    // Protected, wrong password, throws Exception
        //    try
        //    {
        //        wb = WorkbookFactory.Create(
        //                HSSFTestDataSamples.OpenSampleFileStream(xls_prot[0]), "wrong"
        //        );
        //        AssertCloseDoesNotModifyFile(xls_prot[0], wb);
        //        Assert.Fail("Shouldn't be able to open with the wrong password");
        //    }
        //    catch (EncryptedDocumentException e) { }
        //    try
        //    {
        //        wb = WorkbookFactory.Create(
        //                HSSFTestDataSamples.OpenSampleFileStream(xlsx_prot[0]), "wrong"
        //        );
        //        AssertCloseDoesNotModifyFile(xlsx_prot[0], wb);
        //        Assert.Fail("Shouldn't be able to open with the wrong password");
        //    }
        //    catch (EncryptedDocumentException e) { }
        //}
        /**
         * Check that the overloaded file methods which take passwords work properly
         */
        [Test]
        [Ignore("TODO NOT IMPLEMENTED")]
        public void TestCreateWithPasswordFromFile()
        {
            IWorkbook wb;
            // Unprotected, no password given, opens normally
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xls).FullName, null
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            AssertCloseDoesNotModifyFile(xls, wb);

            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xlsx).FullName, null
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            AssertCloseDoesNotModifyFile(xlsx, wb);

            // Unprotected, wrong password, opens normally
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xls).FullName, "wrong"
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            AssertCloseDoesNotModifyFile(xls, wb);

            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xlsx).FullName, "wrong"
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            AssertCloseDoesNotModifyFile(xlsx, wb);

            // Protected, correct password, opens fine
            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xls_prot[0]).FullName, xls_prot[1]
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is HSSFWorkbook);
            AssertCloseDoesNotModifyFile(xls, wb);

            wb = WorkbookFactory.Create(
                    HSSFTestDataSamples.GetSampleFile(xlsx_prot[0]).FullName, xlsx_prot[1]
            );
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb is XSSFWorkbook);
            Assert.IsTrue(wb.NumberOfSheets > 0);
            Assert.IsNotNull(wb.GetSheetAt(0));
            Assert.IsNotNull(wb.GetSheetAt(0).GetRow(0));
            AssertCloseDoesNotModifyFile(xlsx, wb);

            // Protected, wrong password, throws Exception
            try
            {
                wb = WorkbookFactory.Create(
                        HSSFTestDataSamples.GetSampleFile(xls_prot[0]).FullName, "wrong"
                );
                AssertCloseDoesNotModifyFile(xls_prot[0], wb);
                Assert.Fail("Shouldn't be able to open with the wrong password");
            }
            catch (EncryptedDocumentException)
            {
                // expected here
            }
            try
            {
                wb = WorkbookFactory.Create(
                        HSSFTestDataSamples.GetSampleFile(xlsx_prot[0]).FullName, "wrong"
                );
                AssertCloseDoesNotModifyFile(xlsx_prot[0], wb);
                Assert.Fail("Shouldn't be able to open with the wrong password");
            }
            catch (EncryptedDocumentException)
            {
                // expected here
            }
        }
        [Test]
        public void TestEmptyInputStream()
        {
            InputStream emptyStream = new ByteArrayInputStream(new byte[0]);
            try
            {
                WorkbookFactory.Create(emptyStream,true);
                Assert.Fail("Shouldn't be able to create for an empty stream");
            }
            catch (EmptyFileException )
            {
            }
        }
        /**
         * Check that a helpful exception is given on an empty file / stream
         */
        [Test]
        public void TestEmptyFile()
        {
            
            FileInfo emptyFile = TempFile.CreateTempFile("empty", ".poi");
            try
            {
                WorkbookFactory.Create(emptyFile.FullName);
                Assert.Fail("Shouldn't be able to create for an empty file");
            }
            catch (EmptyFileException )
            {
            }
            emptyFile.Delete();
        }
        /**
          * Check that a helpful exception is raised on a non-existing file
          */
        [Test]
        public void TestNonExistantFile()
        {
            FileInfo nonExistantFile = new FileInfo("notExistantFile");
            Assert.IsFalse(nonExistantFile.Exists);
            try
            {
                WorkbookFactory.Create(nonExistantFile.FullName, "password", true);
                Assert.Fail("Should not be able to create for a non-existant file");
            }
            catch (FileNotFoundException)
            {
                // expected
            }
        }

    }

}
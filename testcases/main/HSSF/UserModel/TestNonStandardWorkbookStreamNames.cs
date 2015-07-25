/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using TestCases.HSSF;
    using NUnit.Framework;

    /**
     * Tests for how HSSFWorkbook behaves with XLS files
     *  with a WORKBOOK or BOOK directory entry (instead of 
     *  the more usual, Workbook)
     */
    [TestFixture]
    public class TestNonStandardWorkbookStreamNames
    {

        private String xlsA = "WORKBOOK_in_capitals.xls";
        private String xlsB = "BOOK_in_capitals.xls";

        /**
         * Test that we can Open a file with WORKBOOK
         */
        [Test]
        public void TestOpenWORKBOOK()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(xlsA);

            POIFSFileSystem fs = new POIFSFileSystem(is1);

            // Ensure that we have a WORKBOOK entry
            fs.Root.GetEntry("WORKBOOK");
            // And a summary
            fs.Root.GetEntry("\x0005SummaryInformation");
            Assert.IsTrue(true);

            // But not a Workbook one
            try
            {
                fs.Root.GetEntry("Workbook");
                Assert.Fail();
            }
            catch (FileNotFoundException)
            {

            }

            // Try to Open the workbook
            HSSFWorkbook wb = new HSSFWorkbook(fs);
        }
        /**
        * Test that we can open a file with BOOK
        */
        [Test]
        public void TestOpenBOOK()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(xlsB);

            POIFSFileSystem fs = new POIFSFileSystem(is1);

            // Ensure that we have a BOOK entry
            fs.Root.GetEntry("BOOK");
            Assert.IsTrue(true);

            // But not a Workbook one
            try
            {
                fs.Root.GetEntry("Workbook");
                Assert.Fail();
            }
            catch (FileNotFoundException) { }
            // And not a Summary one
            try
            {
                fs.Root.GetEntry("\005SummaryInformation");
                Assert.Fail();
            }
            catch (FileNotFoundException) { }

            // Try to open the workbook
            HSSFWorkbook wb = new HSSFWorkbook(fs);
        }
        /**
         * Test that when we Write out, we go back to the correct case
         */
        [Test]
        public void TestWrite()
        {
            foreach (String file in new String[] { xlsA, xlsB })
            {
                Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(file);
                POIFSFileSystem fs = new POIFSFileSystem(is1);

                // Open the workbook, not preserving nodes
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                MemoryStream out1 = new MemoryStream();
                wb.Write(out1);

                // Check now
                MemoryStream in1 = new MemoryStream(out1.ToArray());
                POIFSFileSystem fs2 = new POIFSFileSystem(in1);

                // Check that we have the new entries
                fs2.Root.GetEntry("Workbook");
                try
                {
                    fs2.Root.GetEntry("BOOK");
                    Assert.Fail();
                }
                catch (FileNotFoundException) { }
                try
                {
                    fs2.Root.GetEntry("WORKBOOK");
                    Assert.Fail();
                }
                catch (FileNotFoundException) { }

                // And it can be Opened
                HSSFWorkbook wb2 = new HSSFWorkbook(fs2);
            }
        }

        /**
         * Test that when we Write out preserving nodes, we go back to the
         *  correct case
         */
        [Test]
        public void TestWritePreserve()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(xlsA);
            POIFSFileSystem fs = new POIFSFileSystem(is1);

            // Open the workbook, not preserving nodes
            HSSFWorkbook wb = new HSSFWorkbook(fs, true);

            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);

            // Check now
            MemoryStream in1 = new MemoryStream(out1.ToArray());
            POIFSFileSystem fs2 = new POIFSFileSystem(in1);

            // Check that we have the new entries
            fs2.Root.GetEntry("Workbook");
            try
            {
                fs2.Root.GetEntry("WORKBOOK");
                Assert.Fail();
            }
            catch (FileNotFoundException)
            {

            }

            // As we preserved, should also have a few other streams
            fs2.Root.GetEntry("\x0005SummaryInformation");

            // And it can be Opened
            HSSFWorkbook wb2 = new HSSFWorkbook(fs2);
        }
    }
}
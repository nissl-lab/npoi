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
    using NUnit.Framework.Legacy;

    using TestCases.HSSF;
    using NPOI.HPSF;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    /**
     * Old-style setting of POIFS properties doesn't work with POI 3.0.2
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestPOIFSProperties
    {

        private static String title = "Testing POIFS properties";

        [Test]
        public void TestFail()
        {
            MemoryStream out1 = new MemoryStream();
            {
                // read the workbook, adjust the SummaryInformation and write the data to a byte array
                POIFSFileSystem fs = OpenFileSystem();

                HSSFWorkbook wb = new HSSFWorkbook(fs);

                //set POIFS properties After constructing HSSFWorkbook
                //(a piece of code that used to work up to POI 3.0.2)
                SetTitle(fs);

                //save the workbook and read the property
                wb.Write(out1);
                out1.Close();
                wb.Close();
            }

            // process the byte array
            CheckFromByteArray(out1.ToArray());
        }

        [Test]
        public void TestOK()
        {
            MemoryStream out1 = new MemoryStream();
            {
                // read the workbook, adjust the SummaryInformation and write the data to a byte array
                POIFSFileSystem fs = OpenFileSystem();

                //set POIFS properties before constructing HSSFWorkbook
                SetTitle(fs);

                HSSFWorkbook wb = new HSSFWorkbook(fs);

                wb.Write(out1);
                out1.Close();
                wb.Close();
            }

            // process the byte array
            CheckFromByteArray(out1.ToArray());
        }

        private POIFSFileSystem OpenFileSystem()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("Simple.xls");
            POIFSFileSystem fs = new POIFSFileSystem(is1);
            is1.Close();
            return fs;
        }

        private void SetTitle(POIFSFileSystem fs)
        {
            SummaryInformation summary1 = (SummaryInformation) PropertySetFactory.Create(fs.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME));
            ClassicAssert.IsNotNull(summary1);

            summary1.Title = title;
            //write the modified property back to POIFS
            fs.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME).Delete();
            fs.CreateDocument(summary1.ToInputStream(), SummaryInformation.DEFAULT_STREAM_NAME);

            // check that the information was added successfully to the filesystem object
            SummaryInformation summaryCheck = (SummaryInformation) PropertySetFactory.Create(fs.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME));
            ClassicAssert.IsNotNull(summaryCheck);
        }

        private void CheckFromByteArray(byte[] bytes)
        {
            // on some environments in CI we see strange Assert.Failures, let's verify that the size is exactly right
            // this can be removed again After the problem is identified
            // 5120, So 5120 bytes is what we usually get
            int[] allowedSizes = { 5120, 9216 };
            ClassicAssert.IsTrue(System.Array.IndexOf(allowedSizes, bytes.Length) >= 0,
                "Unexpected size " + bytes.Length + ", allowed: " + string.Join(", ", allowedSizes) +
                " Had: " + HexDump.ToHex(bytes));

            POIFSFileSystem fs2 = new POIFSFileSystem(new MemoryStream(bytes));
            SummaryInformation summary2 = (SummaryInformation)PropertySetFactory.Create(fs2.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME));
            ClassicAssert.IsNotNull(summary2);

            ClassicAssert.AreEqual(title, summary2.Title);
            fs2.Close();
        }
    }
}
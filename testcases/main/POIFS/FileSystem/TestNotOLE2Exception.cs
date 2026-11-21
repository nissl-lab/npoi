/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.POIFS.FileSystem
{
    using NPOI.HSSF;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using TestCases.HSSF;

    /// <summary>
    /// Class to test that POIFS complains when given older non-OLE2
    ///  formats. See also <see cref="TestOfficeXMLException"/> for OOXML
    ///  checks
    /// </summary>
    [TestFixture]
    public class TestNotOLE2Exception
    {
        private static Stream OpenXLSSampleStream(string sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
        }
        private static Stream OpenDOCSampleStream(string sampleFileName)
        {
            return POIDataSamples.GetDocumentInstance().OpenResourceAsStream(sampleFileName);
        }

        [Test]
        public void TestRawXMLException()
        {
            Stream in1 = OpenXLSSampleStream("SampleSS.xml");

            try
            {
                new POIFSFileSystem(in1).Close();
                ClassicAssert.Fail("expected exception was not thrown");
            }
            catch(NotOLE2FileException e)
            {
                // expected during successful test
                POITestCase.AssertContains(e.Message, "The supplied data appears to be a raw XML file");
                POITestCase.AssertContains(e.Message, "Formats such as Office 2003 XML");
            }
        }

        [Test]
        public void TestMSWriteException()
        {
            Stream in1 = OpenDOCSampleStream("MSWriteOld.wri");

            try
            {
                new POIFSFileSystem(in1).Close();
                ClassicAssert.Fail("expected exception was not thrown");
            }
            catch(NotOLE2FileException e)
            {
                // expected during successful test
                POITestCase.AssertContains(e.Message, "The supplied data appears to be in the old MS Write");
                POITestCase.AssertContains(e.Message, "doesn't currently support");
            }
        }

        [Test]
        public void TestBiff3Exception()
        {
            Stream in1 = OpenXLSSampleStream("testEXCEL_3.xls");

            try
            {
                new POIFSFileSystem(in1).Close();
                ClassicAssert.Fail("expected exception was not thrown");
            }
            catch(OldExcelFormatException e)
            {
                // expected during successful test
                POITestCase.AssertContains(e.Message, "The supplied data appears to be in BIFF3 format");
                POITestCase.AssertContains(e.Message, "try OldExcelExtractor");
            }
        }

        [Test]
        public void TestBiff4Exception()
        {
            Stream in1 = OpenXLSSampleStream("testEXCEL_4.xls");

            try
            {
                new POIFSFileSystem(in1).Close();
                ClassicAssert.Fail("expected exception was not thrown");
            }
            catch(OldExcelFormatException e)
            {
                // expected during successful test
                POITestCase.AssertContains(e.Message, "The supplied data appears to be in BIFF4 format");
                POITestCase.AssertContains(e.Message, "try OldExcelExtractor");
            }
        }
    }
}

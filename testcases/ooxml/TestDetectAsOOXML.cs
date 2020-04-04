
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

using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NUnit.Framework;
using System.IO;
using TestCases.HSSF;

namespace TestCases.OOXML
{

    /**
     * Class to test that HXF correctly detects OOXML
     *  documents
     */
    [TestFixture]
    public class TestDetectAsOOXML
    {
        [Test]
        public void TestOpensProperly()
        {
            OPCPackage.Open(HSSFTestDataSamples.OpenSampleFileStream("sample.xlsx"));
        }
        [Test]
        public void TestDetectAsPOIFS()
        {
            Stream in1;

            // ooxml file is
            in1 = new PushbackStream(
                    HSSFTestDataSamples.OpenSampleFileStream("SampleSS.xlsx")
            );
            Assert.IsTrue(DocumentFactoryHelper.HasOOXMLHeader(in1));
            in1.Close();

            // xls file isn't
            in1 = new PushbackStream(
                    HSSFTestDataSamples.OpenSampleFileStream("SampleSS.xls")
            );
            Assert.IsFalse(DocumentFactoryHelper.HasOOXMLHeader(in1));
            in1.Close();

            // text file isn't
            in1 = new PushbackStream(
                    HSSFTestDataSamples.OpenSampleFileStream("SampleSS.txt")
            );
            Assert.IsFalse(DocumentFactoryHelper.HasOOXMLHeader(in1));
            in1.Close();
        }
        [Test]
        public void TestFileCorruption()
        {
            // create test InputStream
            byte[] testData = { (byte)1, (byte)2, (byte)3 };
            ByteArrayInputStream testInput = new ByteArrayInputStream(testData);

            // detect header
            InputStream in1 = new PushbackInputStream(testInput, 10);
            Assert.IsFalse(DocumentFactoryHelper.HasOOXMLHeader(in1));
            //noinspection deprecation
            Assert.IsFalse(POIXMLDocument.HasOOXMLHeader(in1));

            // check if InputStream is still intact
            byte[] test = new byte[3];
            Assert.AreEqual(3, in1.Read(test));
            Assert.IsTrue(Arrays.Equals(testData, test));
            Assert.AreEqual(-1, in1.Read());
        }
    }
}



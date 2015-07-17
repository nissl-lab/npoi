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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

namespace TestCases.POIFS.FileSystem
{
    using System;
    using System.Text;
    using System.IO;
    using TestCases.HSSF;
    using NPOI.POIFS.FileSystem;
    using NUnit.Framework;
    using NPOI.Util;


    /**
     * Class to Test that POIFS complains when given an Office 2007 XML document
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestOffice2007XMLException
    {

        private static Stream OpenSampleStream(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
        }
        [Test]
        public void TestXMLException()
        {
            Stream in1 = OpenSampleStream("sample.xlsx");

            try
            {
                new POIFSFileSystem(in1);
                Assert.Fail("expected exception was not thrown");
            }
            catch (OfficeXmlFileException e)
            {
                // expected during successful Test
                Assert.IsTrue(e.Message.IndexOf("The supplied data appears to be in the Office 2007+ XML") > -1);
                Assert.IsTrue(e.Message.IndexOf("You are calling the part of POI that deals with OLE2 Office Documents") > -1);
            }
        }
        [Test]
        public void TestDetectAsPOIFS()
        {

            // ooxml file isn't
            ConfirmIsPOIFS("SampleSS.xlsx", false);

            // xls file is
            ConfirmIsPOIFS("SampleSS.xls", true);

            // text file isn't
            ConfirmIsPOIFS("SampleSS.txt", false);
        }
        private void ConfirmIsPOIFS(String sampleFileName, bool expectedResult)
        {
            Stream in1 = OpenSampleStream(sampleFileName);
            try
            {
                bool actualResult;
                try
                {
                    actualResult = POIFSFileSystem.HasPOIFSHeader(in1);
                }
                catch (IOException ex)
                {
                    throw new RuntimeException(ex);
                }
                Assert.AreEqual(expectedResult, actualResult);
            }
            finally
            {
                in1.Close();
            }

        }
    }
}
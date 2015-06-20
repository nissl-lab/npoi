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
    using System;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HPSF;
    using NPOI.Util;



    /**
     * Tests whether Unicode string can be Read from a
     * DocumentSummaryInformation.
     *
     * @author Rainer Klute (klute@rainer-klute.de)
     * @since 2002-12-09
     * @version $Id: TestUnicode.java 489730 2006-12-22 19:18:16Z bayard $
     */
    [TestFixture]
    public class TestUnicode
    {
        //static string dataDir = @"..\..\..\TestCases\HPSF\data\";
        private static String POI_FS = "TestUnicode.xls";
        private static String[] POI_FILES = new String[]
        {
            "\x0005DocumentSummaryInformation",
        };
//        FileStream data;
        //POIFile[] poiFiles;


        //[SetUp]
        //public void Setup()
        //{
        //    POIDataSamples samples = POIDataSamples.GetHPSFInstance();
        //    data = samples.GetFile(POI_FS);
        //}



        /**
         * Tests the {@link PropertySet} methods. The Test file has two
         * property Set: the first one is a {@link SummaryInformation},
         * the second one is a {@link DocumentSummaryInformation}.
         * 
         * @exception IOException if an I/O exception occurs
         * @exception HPSFException if an HPSF exception occurs
         */
        [Test]
        public void TestPropertySetMethods()
        {
            POIDataSamples samples = POIDataSamples.GetHPSFInstance();
            using (FileStream data = samples.GetFile(POI_FS))
            {

                POIFile poiFile = Util.ReadPOIFiles(data, POI_FILES)[0];
                byte[] b = poiFile.GetBytes();
                PropertySet ps =
                    PropertySetFactory.Create(new ByteArrayInputStream(b));
                Assert.IsTrue(ps.IsDocumentSummaryInformation, "IsDocumentSummaryInformation");
                Assert.AreEqual(ps.SectionCount, 2);
                Section s = (Section)ps.Sections[1];
                Assert.AreEqual(s.GetProperty(1),
                                    CodePageUtil.CP_UTF16);
                Assert.AreEqual(s.GetProperty(2),
                                    -96070278);
                Assert.AreEqual(s.GetProperty(3),
                                    "MCon_Info zu Office bei Schreiner");
                Assert.AreEqual(s.GetProperty(4),
                                    "petrovitsch@schreiner-online.de");
                Assert.AreEqual(s.GetProperty(5),
                                    "Petrovitsch, Wilhelm");
            }
        }
    }
}
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
     * Test case for OLE2 files with empty properties. An empty property's type
     * is {@link Variant#VT_EMPTY}.
     *
     * @author Rainer Klute <a
     * href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
     * @since 2003-07-25
     * @version $Id: TestEmptyProperties.java 489730 2006-12-22 19:18:16Z bayard $
     */
    [TestFixture]
    public class TestEmptyProperties
    {
       // static string dataDir = @"..\..\..\TestCases\HPSF\data\";
        /**
         * This Test file's summary information stream Contains some empty
         * properties.
         */
        static String POI_FS = "TestCorel.shw";

        static String[] POI_FILES = new String[]
        {
            "PerfectOffice_MAIN",
            "\x0005SummaryInformation",
            "Main"
        };

        POIFile[] poiFiles;

        [SetUp]
        public void SetUp()
        {
            POIDataSamples samples = POIDataSamples.GetHPSFInstance();
            Stream data = samples.OpenResourceAsStream(POI_FS);;
            poiFiles = Util.ReadPOIFiles(data);
        }
        /**
         * Checks the names of the files in the POI filesystem. They
         * are expected to be in a certain order.
         */
        [Test]
        public void TestReadFiles()
        {
            String[] expected = POI_FILES;
            for (int i = 0; i < expected.Length; i++)
                Assert.AreEqual(poiFiles[i].GetName(), expected[i]);
        }

        /**
         * Tests whether property Sets can be Created from the POI
         * files in the POI file system. This Test case expects the first
         * file to be a {@link SummaryInformation}, the second file to be
         * a {@link DocumentSummaryInformation} and the rest to be no
         * property Sets. In the latter cases a {@link
         * NoPropertySetStreamException} will be thrown when trying to
         * Create a {@link PropertySet}.
         * 
         * @exception IOException if an I/O exception occurs.
         * 
         * @exception UnsupportedEncodingException if a character encoding is not
         * supported.
         */
        [Test]
        public void TestCreatePropertySets()
        {
            Type[] expected = new Type[]
            {
                typeof(NoPropertySetStreamException),
                typeof(SummaryInformation),
                typeof(NoPropertySetStreamException)
            };
            for (int i = 0; i < expected.Length; i++)
            {
                InputStream in1 = new ByteArrayInputStream(poiFiles[i].GetBytes());
                Object o;
                try
                {
                    o = PropertySetFactory.Create(in1);
                }
                catch (NoPropertySetStreamException ex)
                {
                    o = ex;
                }
                catch (MarkUnsupportedException ex)
                {
                    o = ex;
                }
                in1.Close();
                Assert.AreEqual(o.GetType(), expected[i]);
            }
        }



        /**
         * Tests the {@link PropertySet} methods. The Test file has two
         * property Sets: the first one is a {@link SummaryInformation},
         * the second one is a {@link DocumentSummaryInformation}.
         * 
         * @exception IOException if an I/O exception occurs
         * @exception HPSFException if an HPSF operation fails
         */
        [Test]
        public void TestPropertySetMethods()
        {
            byte[] b = poiFiles[1].GetBytes();
            PropertySet ps =
                PropertySetFactory.Create(new ByteArrayInputStream(b));
            SummaryInformation s = (SummaryInformation)ps;
            Assert.IsNull(s.Title);
            Assert.IsNull(s.Subject);
            Assert.IsNotNull(s.Author);
            Assert.IsNull(s.Keywords);
            Assert.IsNull(s.Comments);
            Assert.IsNotNull(s.Template);
            Assert.IsNotNull(s.LastAuthor);
            Assert.IsNotNull(s.RevNumber);
            Assert.AreEqual(s.EditTime, 0);
            Assert.IsNull(s.LastPrinted);
            Assert.IsNull(s.CreateDateTime);
            Assert.IsNull(s.LastSaveDateTime);
            Assert.AreEqual(s.PageCount, 0);
            Assert.AreEqual(s.WordCount, 0);
            Assert.AreEqual(s.CharCount, 0);
            Assert.IsNull(s.Thumbnail);
            Assert.IsNull(s.ApplicationName);
        }

    }
}
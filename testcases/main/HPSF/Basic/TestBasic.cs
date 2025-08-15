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
    using NPOI.HPSF;
    using NPOI.HPSF.Wellknown;
    using NPOI.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;


    /**
     * Tests the basic HPSF functionality.
     *
     * @author Rainer Klute (klute@rainer-klute.de)
     * @since 2002-07-20
     * @version $Id: TestBasic.java 619848 2008-02-08 11:55:43Z klute $
     */
    [TestFixture]
    public class TestBasic
    {
        //static string dataDir = @"..\..\..\TestCases\HPSF\data\";
        private static POIDataSamples samples = POIDataSamples.GetHPSFInstance();

        private static String[] POI_FILES = {
            "\x0005SummaryInformation",
            "\x0005DocumentSummaryInformation",
            "WordDocument",
            "\x0001CompObj",
            "1Table"
        };
        private static int BYTE_ORDER   = 0xfffe;
        private static int FORMAT       = 0x0000;
        private static int OS_VERSION   = 0x00020A04;
        private static ClassID CLASS_ID = new ClassID("{00000000-0000-0000-0000-000000000000}");
        private static int[] SECTION_COUNT = {1, 2};
        private static bool[] IS_SUMMARY_INFORMATION = {true, false};
        private static bool[] IS_DOCUMENT_SUMMARY_INFORMATION = {false, true};

        List<POIFile> poiFiles;



        
        /// <summary>
        /// Read a the test file from the "data" directory.
        /// </summary>
        /// <exception cref="FileNotFoundException">if the file to be read does not exist.</exception>
        /// <exception cref="IOException">if any other I/O exception occurs.</exception>
        [SetUp]
        public void SetUp()
        {
            FileStream data = samples.GetFile("TestGermanWord90.doc");
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
                ClassicAssert.AreEqual(poiFiles[i].GetName(), expected[i]);
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
                typeof(SummaryInformation),
                typeof(DocumentSummaryInformation),
                typeof(NoPropertySetStreamException),
                typeof(NoPropertySetStreamException),
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
                ClassicAssert.AreEqual(expected[i], o.GetType());
            }
        }



        /**
         * Tests the {@link PropertySet} methods. The Test file has two
         * property Sets: the first one is a {@link SummaryInformation},
         * the second one is a {@link DocumentSummaryInformation}.
         * 
         * @exception IOException if an I/O exception occurs
         * @exception HPSFException if any HPSF exception occurs
         */
        [Test]
        public void TestPropertySetMethods()
        {
            /* Loop over the two property Sets. */
            for (int i = 0; i < 2; i++)
            {
                byte[] b = poiFiles[i].GetBytes();
                PropertySet ps = PropertySetFactory.Create(new ByteArrayInputStream(b));
                ClassicAssert.AreEqual(ps.ByteOrder, BYTE_ORDER);
                ClassicAssert.AreEqual(ps.Format, FORMAT);
                ClassicAssert.AreEqual(ps.OSVersion, OS_VERSION);
                ClassicAssert.AreEqual(CLASS_ID, ps.ClassID);
                ClassicAssert.AreEqual(SECTION_COUNT[i], ps.SectionCount);
                ClassicAssert.AreEqual(IS_SUMMARY_INFORMATION[i], ps.IsSummaryInformation);
                ClassicAssert.AreEqual(IS_DOCUMENT_SUMMARY_INFORMATION[i], ps.IsDocumentSummaryInformation);
            }
        }



        /**
         * Tests the {@link Section} methods. The Test file has two
         * property Sets: the first one is a {@link SummaryInformation},
         * the second one is a {@link DocumentSummaryInformation}.
         * 
         * @exception IOException if an I/O exception occurs
         * @exception HPSFException if any HPSF exception occurs
         */
        [Test]
        public void TestSectionMethods()
        {
            InputStream is1 = new ByteArrayInputStream(poiFiles[0].GetBytes());
            SummaryInformation si = (SummaryInformation)PropertySetFactory.Create(is1);
            IList sections = si.Sections;
            Section s = (Section)sections[0];
            
            ClassicAssert.AreEqual(s.FormatID, SectionIDMap.SUMMARY_INFORMATION_ID);
            ClassicAssert.IsNotNull(s.Properties);
            ClassicAssert.AreEqual(17, s.PropertyCount);
            ClassicAssert.AreEqual("Titel", s.GetProperty(2));
            ClassicAssert.AreEqual(1764, s.Size);
        }

        [Test]
        public void Bug52117LastPrinted()
        {
            FileStream f = samples.GetFile("TestBug52117.doc");
            POIFile poiFile = Util.ReadPOIFiles(f, new String[]{POI_FILES[0]})[0];
            InputStream in1 = new ByteArrayInputStream(poiFile.GetBytes());
            SummaryInformation si = (SummaryInformation)PropertySetFactory.Create(in1);
            var lastPrinted = si.LastPrinted;
            long editTime = si.EditTime;
            ClassicAssert.IsTrue(Filetime.IsUndefined(lastPrinted));
            ClassicAssert.AreEqual(1800000000L, editTime);
        }
    }
}
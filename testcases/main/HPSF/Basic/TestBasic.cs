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
    using System.Text;
    using System.Collections;
    using NUnit.Framework;
    using NPOI.HPSF;
    using NPOI.HPSF.Wellknown;
    using NPOI.Util;


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
        static String POI_FS = "TestGermanWord90.doc";
        static String[] POI_FILES = new String[]
        {
            "\x0005SummaryInformation",
            "\x0005DocumentSummaryInformation",
            "WordDocument",
            "\x0001CompObj",
            "1Table"
        };
        static int BYTE_ORDER = 0xfffe;
        static int FORMAT = 0x0000;
        static int OS_VERSION = 0x00020A04;
        static byte[] CLASS_ID =
        {
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00
        };
        static int[] SECTION_COUNT = { 1, 2 };
        static bool[] IS_SUMMARY_INFORMATION = { true, false };
        static bool[] IS_DOCUMENT_SUMMARY_INFORMATION = { false, true };

        POIFile[] poiFiles;



        /**
         * Test case constructor.
         * 
         * @param name The Test case's name.
         */
        public TestBasic()
        {
            //FileStream data =File.OpenRead(dataDir+POI_FS);
            POIDataSamples samples = POIDataSamples.GetHPSFInstance();
            using (Stream data = (Stream)samples.OpenResourceAsStream(POI_FS))
            {
                poiFiles = Util.ReadPOIFiles(data);
            }
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
                Assert.AreEqual(expected[i], o.GetType());
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
                Assert.AreEqual(ps.ByteOrder, BYTE_ORDER);
                Assert.AreEqual(ps.Format, FORMAT);
                Assert.AreEqual(ps.OSVersion, OS_VERSION);
                CollectionAssert.AreEqual(CLASS_ID, ps.ClassID.Bytes);
                Assert.AreEqual(SECTION_COUNT[i], ps.SectionCount);
                Assert.AreEqual(IS_SUMMARY_INFORMATION[i], ps.IsSummaryInformation);
                Assert.AreEqual(IS_DOCUMENT_SUMMARY_INFORMATION[i], ps.IsDocumentSummaryInformation);
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
            SummaryInformation si = (SummaryInformation)
                PropertySetFactory.Create(new ByteArrayInputStream
                    (poiFiles[0].GetBytes()));
            IList sections = si.Sections;
            Section s = (Section)sections[0];
            Assert.IsTrue(Arrays.Equals
                (s.FormatID.Bytes, SectionIDMap.SUMMARY_INFORMATION_ID));
            Assert.IsNotNull(s.Properties);
            Assert.AreEqual(17, s.PropertyCount);
            Assert.AreEqual("Titel", s.GetProperty(2));
            //Assert.assertEquals(1764, s.getSize());
            Assert.AreEqual(1764, s.Size);
        }

    }
}
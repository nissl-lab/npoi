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
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.HPSF;
    using NPOI.Util;
    using System.Collections.Generic;
    using NPOI.POIFS.FileSystem;
    using System.Text.RegularExpressions;
    using System.Linq;


    /**
     * Tests some HPSF functionality by Reading all property Sets from all files
     * in the "data" directory. If you want to ensure HPSF can deal with a certain
     * OLE2 file, just Add it to the "data" directory and run this Test case.
     * 
     * @author Rainer Klute (klute@rainer-klute.de)
     * @since 2008-02-08
     * @version $Id: TestBasic.java 489730 2006-12-22 19:18:16Z bayard $
     */
    [TestFixture]
    [TestFixtureSource(typeof(AllHPSFTestFile), nameof(AllHPSFTestFile.Files))]
    public class TestReadAllFiles
    {
        public FileInfo file;
        public TestReadAllFiles(FileInfo fileinfo)
        {
            this.file = fileinfo;
        }
        private static String[] excludes = new String[] {
            //"TestInvertedClassID.doc", //failed with 'MacRoman' is not a supported encoding name.
            //"TestBug52372.doc", //failed with 'MacRoman' is not a supported encoding name.
        };
        POIDataSamples _samples = POIDataSamples.GetHPSFInstance();
        /**
         * This Test methods Reads all property Set streams from all POI
         * filesystems in the "data" directory.
         */

        [Test]
        public void TestReadAllFiles1()
        {
            foreach (POIFile pf in Util.ReadPropertySets(file))
            {
                InputStream in1 =
                    new ByteArrayInputStream(pf.GetBytes());
                try
                {
                    PropertySet a = PropertySetFactory.Create(in1);
                }
                finally
                {
                    in1.Close();
                }
            }
        }

        /**
         * Returns true if the file should be checked, false if it should be excluded.
         *
         * @param f
         * @return
         */
        public static bool checkExclude(string f)
        {
            foreach (String exclude in excludes)
            {
                if (f.EndsWith(exclude))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// <para>
        /// This test method does a write and read back test with all POI
        /// filesystems in the "data" directory by performing the following
        /// actions for each file:
        /// </para>
        /// <para>
        /// <list type="bullet">
        /// <item><description>Read its property Set streams.</description></item>
        /// <item><description>Create a new POI filesystem containing the origin file's property Set streams.</description></item>
        /// <item><description>Read the property Set streams from the POI filesystem just created.</description></item>
        /// <item><description>Compare each property Set stream with the corresponding one from
        /// the origin file and check whether they are equal.</description></item>
        /// </list>
        /// </para>
        /// </summary>
        [Test]
        public void Recreate() {
            /* Read the POI filesystem's property Set streams: */
            Dictionary<String,PropertySet> psMap = new Dictionary<String,PropertySet>();
        
            /* Create a new POI filesystem containing the origin file's
             * property Set streams: */
            
            POIFSFileSystem poiFs = new POIFSFileSystem();
            foreach (POIFile poifile in Util.ReadPropertySets(file))
            {
                InputStream in1 = new ByteArrayInputStream(poifile.GetBytes());
                PropertySet psIn = PropertySetFactory.Create(in1);
                psMap[poifile.GetName()] = psIn;
                //bos.Seek(0, SeekOrigin.Begin);
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                psIn.Write(bos);
                poiFs.CreateDocument(new MemoryStream(bos.ToByteArray()), poifile.GetName());
            }

            /* Read the property Set streams from the POI filesystem just
             * created. */
            foreach (KeyValuePair<String,PropertySet> me in psMap)
            {
                PropertySet ps1 = me.Value;
                PropertySet ps2 = PropertySetFactory.Create(poiFs.Root, me.Key);
                ClassicAssert.IsNotNull(ps2);
            
                /* Compare the property Set stream with the corresponding one
                 * from the origin file and check whether they are equal. */
            
                // Because of missing 0-paddings in the original input files, the bytes might differ.
                // This fixes the comparison
                
                string pattern = "(?m)(\\s+$|(size|offset): [0-9]+)";

                string ps1str = Regex.Replace(ps1.ToString().Replace(" 00", "   ").Replace(".", " "), pattern, "");
                string ps2str = Regex.Replace(ps2.ToString().Replace(" 00", "   ").Replace(".", " "), pattern, "");
            
                ClassicAssert.AreEqual(ps1str, ps2str, "Equality for file " + file.Name);
            }
            poiFs.Close();
        }
    
        /// <summary>
        /// This test method checks whether DocumentSummary information streams
        /// can be read. This is done by opening all "Test*" files in the 'poifs' directrory
        /// pointed to by the "POI.testdata.path" system property, trying to extract
        /// the document summary information stream in the root directory and calling
        /// its Get... methods.
        /// </summary>
        /// <exception cref="Exception"></exception>
        [Test]
        public void ReadDocumentSummaryInformation()
        {
            /* Read a test document <em>doc</em> into a POI filesystem. */
            NPOIFSFileSystem poifs = new NPOIFSFileSystem(file, true);
            try
            {
                DirectoryEntry dir = poifs.Root;
                /*
                 * If there is a document summry information stream, read it from
                 * the POI filesystem.
                 */
                if (dir.HasEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME))
                {
                    DocumentSummaryInformation dsi = TestWriteWellKnown.GetDocumentSummaryInformation(poifs);
    
                    /* Execute the Get... methods. */
                    var _ = dsi.ByteCount;
                    _ = dsi.ByteOrder;
                    var s = dsi.Category;
                    s = dsi.Company;
                    var a = dsi.CustomProperties;
                    // FIXME dsi.Docparts;
                    // FIXME dsi.HeadingPair;
                    _ = dsi.HiddenCount;
                    _ = dsi.LineCount;
                    var b = dsi.LinksDirty;
                    s = dsi.Manager;
                    _ = dsi.MMClipCount;
                    _ = dsi.NoteCount;
                    _ = dsi.ParCount;
                    s = dsi.PresentationFormat;
                    b = dsi.Scale;
                    _ = dsi.SlideCount;
                }
            }
            finally
            {
                poifs.Close();
            }
        }
    
        /// <summary>
        /// Tests the simplified custom properties by reading them from the
        /// available test files.
        /// </summary>
        /// <exception cref="Throwable">if anything goes wrong.</exception>
        [Test]
        public void ReadCustomPropertiesFromFiles()
        {
            /* Read a test document <em>doc</em> into a POI filesystem. */
            NPOIFSFileSystem poifs = new NPOIFSFileSystem(file);
            try
            {
                /*
                 * If there is a document summry information stream, read it from
                 * the POI filesystem, else create a new one.
                 */
                DocumentSummaryInformation dsi = TestWriteWellKnown.GetDocumentSummaryInformation(poifs);
                if (dsi == null) 
                {
                    dsi = PropertySetFactory.NewDocumentSummaryInformation();
                }
                CustomProperties cps = dsi.CustomProperties;

                if (cps == null)
                {
                    /* The document does not have custom properties. */
                    return;
                }

                foreach (CustomProperty cp in cps.Properties())
                {
                    ClassicAssert.IsNotNull(cp.Name);
                    ClassicAssert.IsNotNull(cp.Value);
                }
            }
            finally
            {
                poifs.Close();
            }
        }
    }

    public class AllHPSFTestFile
    {
        public static IEnumerable Files
        {
            get
            {
                POIDataSamples _samples = POIDataSamples.GetHPSFInstance();
                string[] files = _samples.GetFiles();
                return files.Where(x=>TestReadAllFiles.checkExclude(x)).Select(f=>new FileInfo(f));
            }
        }
    }
}
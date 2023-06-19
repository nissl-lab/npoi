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
    using System.Collections;
    using System.IO;
    using System.Text;

    using NPOI.HPSF;
    using NPOI.HPSF.Wellknown;
    using NPOI.POIFS.EventFileSystem;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    using NUnit.Framework;
    using NUnit.Framework.Constraints;

    /**
     * Tests HPSF's writing functionality.
     *
     * @author Rainer Klute (klute@rainer-klute.de)
     * @since 2003-02-07
     * @version $Id: TestWrite.java 489730 2006-12-22 19:18:16Z bayard $
     */
    [TestFixture]
    public class TestWrite
    {
        static String POI_FS = "TestHPSFWritingFunctionality.doc";
        private static POIDataSamples _samples = POIDataSamples.GetHPSFInstance();

        static int[] SECTION_COUNT = { 1, 2 };
        static bool[] IS_SUMMARY_INFORMATION = { true, false };
        static bool[] IS_DOCUMENT_SUMMARY_INFORMATION = { false, true };

        String IMPROPER_DEFAULT_CHARSET_MESSAGE =
            "Your default character Set is " + GetDefaultCharSetName() +
            ". However, this Testcase must be run in an environment " +
            "with a default character Set supporting at least " +
            "8-bit-characters. You can achieve this by Setting the " +
            "LANG environment variable to a proper value, e.g. " +
            "\"de_DE\".";

        [SetUp]
        public void SetUp()
        {
            VariantSupport.IsLogUnsupportedTypes = (false);
        }



        /**
         * Writes an empty property Set to a POIFS and Reads it back
         * in.
         * 
         * @exception IOException if an I/O exception occurs
         */
        [Test]
        public void TestWithoutAFormatID()
        {
            FileInfo fi = TempFile.CreateTempFile(POI_FS, ".doc");
            using (FileStream file = new FileStream(fi.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                //FileStream filename = File.OpenRead(dataDir + POI_FS);
                //filename.deleteOnExit();

                /* Create a mutable property Set with a section that does not have the
                 * formatID Set: */
                FileStream out1 = file;
                POIFSFileSystem poiFs = new POIFSFileSystem();
                MutablePropertySet ps = new MutablePropertySet();
                ps.ClearSections();
                ps.AddSection(new MutableSection());

                /* Write it to a POIFS and the latter to disk: */
                try
                {
                    MemoryStream psStream = new MemoryStream();
                    ps.Write(psStream);
                    psStream.Close();
                    byte[] streamData = psStream.ToArray();
                    poiFs.CreateDocument(new MemoryStream(streamData),
                                         SummaryInformation.DEFAULT_STREAM_NAME);
                    poiFs.WriteFileSystem(out1);
                    out1.Close();
                    Assert.Fail("Should have thrown a NoFormatIDException.");
                }
                catch (Exception ex)
                {
                    Assert.IsTrue(ex is NoFormatIDException);
                }
                finally
                {
                    poiFs.Close();
                    out1.Close();
                }
            }
            try
            {
                File.Delete(fi.FullName);
            }
            catch
            {
            }
        }



        /**
         * Writes an empty property Set to a POIFS and Reads it back
         * in.
         * 
         * @exception IOException if an I/O exception occurs
         * @exception UnsupportedVariantTypeException if HPSF does not yet support
         * a variant type to be written
         */
        [Test]
        public void TestWriteEmptyPropertySet()
        {
            FileInfo fi = TempFile.CreateTempFile(POI_FS, ".doc");
            using (FileStream file = new FileStream(fi.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                //filename.deleteOnExit();

                /* Create a mutable property Set and Write it to a POIFS: */
                FileStream out1 = file;
                POIFSFileSystem poiFs = new POIFSFileSystem();
                MutablePropertySet ps = new MutablePropertySet();
                MutableSection s = (MutableSection)ps.Sections[0];
                s.SetFormatID(SectionIDMap.SUMMARY_INFORMATION_ID);

                MemoryStream psStream = new MemoryStream();
                ps.Write(psStream);
                psStream.Close();
                byte[] streamData = psStream.ToArray();
                poiFs.CreateDocument(new MemoryStream(streamData),
                                     SummaryInformation.DEFAULT_STREAM_NAME);
                poiFs.WriteFileSystem(out1);
                //out1.Close();
                file.Position = 0;
                /* Read the POIFS: */
                POIFSReader reader3 = new POIFSReader();
                reader3.StreamReaded += new POIFSReaderEventHandler(reader3_StreamReaded);
                reader3.Read(file);

                poiFs.Close();
                file.Close();
                //File.Delete(dataDir + POI_FS);
            }
            try
            {
                File.Delete(fi.FullName);
            }
            catch
            {
            }
        }

        void reader3_StreamReaded(object sender, POIFSReaderEventArgs e)
        {
            try
            {
                PropertySetFactory.Create(e.Stream);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }


        /* Read the POIFS: */
        static PropertySet[] psa = new PropertySet[1];
        /**
         * Writes a simple property Set with a SummaryInformation section to a
         * POIFS and Reads it back in.
         * 
         * @exception IOException if an I/O exception occurs
         * @exception UnsupportedVariantTypeException if HPSF does not yet support
         * a variant type to be written
         */
        [Test]
        public void TestWriteSimplePropertySet()
        {
            String AUTHOR = "Rainer Klute";
            String TITLE = "Test Document";

            FileInfo fi = TempFile.CreateTempFile(POI_FS, ".doc");
            FileStream file = new FileStream(fi.FullName, FileMode.Open, FileAccess.ReadWrite);

            FileStream out1 = file;
            POIFSFileSystem poiFs = new POIFSFileSystem();

            MutablePropertySet ps = new MutablePropertySet();
            MutableSection si = new MutableSection();
            si.SetFormatID(SectionIDMap.SUMMARY_INFORMATION_ID);
            ps.Sections[0] = si;

            MutableProperty p = new MutableProperty();
            p.ID = PropertyIDMap.PID_AUTHOR;
            p.Type = Variant.VT_LPWSTR;
            p.Value = AUTHOR;
            si.SetProperty(p);
            si.SetProperty(PropertyIDMap.PID_TITLE, Variant.VT_LPSTR, TITLE);

            poiFs.CreateDocument(ps.ToInputStream(),
                                 SummaryInformation.DEFAULT_STREAM_NAME);
            poiFs.WriteFileSystem(out1);
            poiFs.Close();
            //out1.Close();
            file.Position = 0;

            POIFSReader reader1 = new POIFSReader();
            //reader1.StreamReaded += new POIFSReaderEventHandler(reader1_StreamReaded);
            POIFSReaderListener1 psl = new POIFSReaderListener1();
            reader1.RegisterListener(psl);
            reader1.Read(file);
            Assert.IsNotNull(psa[0]);
            Assert.IsTrue(psa[0].IsSummaryInformation);

            Section s = (Section)(psa[0].Sections[0]);
            Object p1 = s.GetProperty(PropertyIDMap.PID_AUTHOR);
            Object p2 = s.GetProperty(PropertyIDMap.PID_TITLE);
            Assert.AreEqual(AUTHOR, p1);
            Assert.AreEqual(TITLE, p2);
            file.Close();
            try
            {
                File.Delete(fi.FullName);
            }
            catch
            {
            }
        }
        private class POIFSReaderListener1 : POIFSReaderListener
        {
            public void ProcessPOIFSReaderEvent(POIFSReaderEvent e)
            {
                try
                {
                    psa[0] = PropertySetFactory.Create(e.Stream);
                }
                catch (Exception ex)
                {
                    Assert.Fail(ex.Message);
                }
            }
        }

        /**
         * Writes a simple property Set with two sections to a POIFS and Reads it
         * back in.
         * 
         * @exception IOException if an I/O exception occurs
         * @exception WritingNotSupportedException if HPSF does not yet support
         * a variant type to be written
         */
        [Test]
        public void TestWriteTwoSections()
        {
            String STREAM_NAME = "PropertySetStream";
            String SECTION1 = "Section 1";
            String SECTION2 = "Section 2";

            FileInfo fi = TempFile.CreateTempFile(POI_FS, ".doc");
            FileStream file = new FileStream(fi.FullName, FileMode.Open, FileAccess.ReadWrite);
            //filename.deleteOnExit();
            FileStream out1 = file;

            POIFSFileSystem poiFs = new POIFSFileSystem();
            MutablePropertySet ps = new MutablePropertySet();
            ps.ClearSections();

            ClassID formatID = new ClassID();
            formatID.Bytes = new byte[]{0, 1,  2,  3,  4,  5,  6,  7,
                                     8, 9, 10, 11, 12, 13, 14, 15};
            MutableSection s1 = new MutableSection();
            s1.SetFormatID(formatID);
            s1.SetProperty(2, SECTION1);
            ps.AddSection(s1);

            MutableSection s2 = new MutableSection();
            s2.SetFormatID(formatID);
            s2.SetProperty(2, SECTION2);
            ps.AddSection(s2);

            poiFs.CreateDocument(ps.ToInputStream(), STREAM_NAME);
            poiFs.WriteFileSystem(out1);
            poiFs.Close();
            //out1.Close();

            /* Read the POIFS: */
            psa = new PropertySet[1];
            POIFSReader reader2 = new POIFSReader();
            //reader2.StreamReaded += new POIFSReaderEventHandler(reader2_StreamReaded);
            POIFSReaderListener2 prl = new POIFSReaderListener2();
            reader2.RegisterListener(prl);
            reader2.Read(file);
            Assert.IsNotNull(psa[0]);
            Section s = (Section)(psa[0].Sections[0]);
            Assert.AreEqual(s.FormatID, formatID);
            Object p = s.GetProperty(2);
            Assert.AreEqual(SECTION1, p);
            s = (Section)(psa[0].Sections[1]);
            p = s.GetProperty(2);
            Assert.AreEqual(SECTION2, p);

            file.Close();
            //File.Delete(dataDir + POI_FS);
            try
            {
                File.Delete(fi.FullName);
            }
            catch
            {
            }
        }
        private class POIFSReaderListener2 : POIFSReaderListener
        {
            #region POIFSReaderListener ³ÉÔ±

            public void ProcessPOIFSReaderEvent(POIFSReaderEvent evt)
            {
                try
                    {
                        psa[0] = PropertySetFactory.Create(evt.Stream);
                    }
                    catch (Exception ex)
                    {
                        throw new RuntimeException(ex.Message);
                        /* FIXME (2): Replace the previous line by the following
                         * one once we no longer need JDK 1.3 compatibility. */
                        // throw new RuntimeException(ex);
                    }
            }

            #endregion
        }
        



        private static int CODEPAGE_DEFAULT = -1;
        private static int CODEPAGE_1252 = 1252;
        private static int CODEPAGE_UTF8 = CodePageUtil.CP_UTF8;
        private static int CODEPAGE_UTF16 = CodePageUtil.CP_UTF16;



        /**
         * Writes and Reads back various variant types and checks whether the
         * stuff that has been Read back Equals the stuff that was written.
         */
        [Test]
        public void TestVariantTypes()
        {
            int codepage = CODEPAGE_DEFAULT;
            if (!hasProperDefaultCharSet())
            {
                Console.Error.WriteLine(IMPROPER_DEFAULT_CHARSET_MESSAGE +
                    " This Testcase is skipped.");
                return;
            }

            check(Variant.VT_EMPTY, null, codepage);
            check(Variant.VT_BOOL, true, codepage);
            check(Variant.VT_BOOL, false, codepage);
            check(Variant.VT_CF, new byte[] { 8, 0, 0, 0, 1, 0, 0, 0, 1, 2, 3,
                    4 }, codepage);
            check(Variant.VT_I4, 27, codepage);
            check(Variant.VT_I8, (long)28, codepage);
            check(Variant.VT_R8, 29.0, codepage);
            check(Variant.VT_I4, -27, codepage);
            check(Variant.VT_I8, (long)-28, codepage);
            check(Variant.VT_R8, -29.0, codepage);
            check(Variant.VT_FILETIME, new DateTime(1984, 5, 16, 8, 23, 15), codepage);
            check(Variant.VT_I4, int.MaxValue, codepage);
            check(Variant.VT_I4, int.MinValue, codepage);
            check(Variant.VT_I8, long.MaxValue, codepage);
            check(Variant.VT_I8, long.MinValue, codepage);
            check(Variant.VT_R8, Double.MaxValue, codepage);
            check(Variant.VT_R8, Double.MinValue, codepage);

            check(Variant.VT_LPSTR,
                  "", codepage);
            check(Variant.VT_LPSTR,
                  "\u00e4", codepage);
            check(Variant.VT_LPSTR,
                  "\u00e4\u00f6", codepage);
            check(Variant.VT_LPSTR,
                  "\u00e4\u00f6\u00fc", codepage);
            check(Variant.VT_LPSTR,
                  "\u00e4\u00f6\u00fc\u00df", codepage);
            check(Variant.VT_LPSTR,
                  "\u00e4\u00f6\u00fc\u00df\u00c4", codepage);
            check(Variant.VT_LPSTR,
                  "\u00e4\u00f6\u00fc\u00df\u00c4\u00d6", codepage);
            check(Variant.VT_LPSTR,
                  "\u00e4\u00f6\u00fc\u00df\u00c4\u00d6\u00dc", codepage);

            check(Variant.VT_LPWSTR,
                  "", codepage);
            check(Variant.VT_LPWSTR,
                  "\u00e4", codepage);
            check(Variant.VT_LPWSTR,
                  "\u00e4\u00f6", codepage);
            check(Variant.VT_LPWSTR,
                  "\u00e4\u00f6\u00fc", codepage);
            check(Variant.VT_LPWSTR,
                  "\u00e4\u00f6\u00fc\u00df", codepage);
            check(Variant.VT_LPWSTR,
                  "\u00e4\u00f6\u00fc\u00df\u00c4", codepage);
            check(Variant.VT_LPWSTR,
                  "\u00e4\u00f6\u00fc\u00df\u00c4\u00d6", codepage);
            check(Variant.VT_LPWSTR,
                  "\u00e4\u00f6\u00fc\u00df\u00c4\u00d6\u00dc", codepage);
        }



        /**
         * Writes and Reads back strings using several different codepages and
         * checks whether the stuff that has been Read back Equals the stuff that
         * was written.
         */
        [Test]
        public void TestCodepages()
        {
            //Exception thr = null;
            int[] validCodepages = new int[] { CODEPAGE_DEFAULT, CODEPAGE_UTF8, CODEPAGE_UTF16, CODEPAGE_1252 };
            for (int i = 0; i < validCodepages.Length; i++)
            {
                int cp = validCodepages[i];
                if (cp == -1 && !hasProperDefaultCharSet())
                {
                    Console.Error.WriteLine(IMPROPER_DEFAULT_CHARSET_MESSAGE +
                         " This Testcase is skipped for the default codepage.");
                    continue;
                }

                long t = cp == CODEPAGE_UTF16 ? Variant.VT_LPWSTR
                                                    : Variant.VT_LPSTR;
                try
                {
                    check(t, "", cp);
                    check(t, "\u00e4", cp);
                    check(t, "\u00e4\u00f6", cp);
                    check(t, "\u00e4\u00f6\u00fc", cp);
                    check(t, "\u00e4\u00f6\u00fc\u00c4", cp);
                    check(t, "\u00e4\u00f6\u00fc\u00c4\u00d6", cp);
                    check(t, "\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc", cp);
                    check(t, "\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df", cp);
                    if (cp == CodePageUtil.CP_UTF16 || cp == CodePageUtil.CP_UTF8)
                        check(t, "\u79D1\u5B78", cp);
                }
                catch (Exception ex)
                {
                    Assert.Fail(ex.Message +
                         " with codepage " + cp);
                }
            }

            int[] invalidCodepages = new int[] { 0, 1, 2, 4711, 815 };
            for (int i = 0; i < invalidCodepages.Length; i++)
            {
                int cp = invalidCodepages[i];
                long type = cp == CODEPAGE_UTF16 ? Variant.VT_LPWSTR
                                                       : Variant.VT_LPSTR;

                try
                {
                    check(type, "", cp);
                    check(type, "\u00e4", cp);
                    check(type, "\u00e4\u00f6", cp);
                    check(type, "\u00e4\u00f6\u00fc", cp);
                    check(type, "\u00e4\u00f6\u00fc\u00c4", cp);
                    check(type, "\u00e4\u00f6\u00fc\u00c4\u00d6", cp);
                    check(type, "\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc", cp);
                    check(type, "\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df", cp);

                    Assert.Fail("UnsupportedEncodingException for codepage " + cp +
                     " expected.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }

        }



        /**
         * Tests whether writing 8-bit characters to a Unicode property
         * succeeds.
         */
        [Test]
        public void TestUnicodeWrite8Bit()
        {
            String TITLE = "This is a sample title";
            MutablePropertySet mps = new MutablePropertySet();
            MutableSection ms = (MutableSection)mps.Sections[0];
            ms.SetFormatID(SectionIDMap.SUMMARY_INFORMATION_ID);
            MutableProperty p = new MutableProperty();
            p.ID = PropertyIDMap.PID_TITLE;
            p.Type = Variant.VT_LPSTR;
            p.Value = TITLE;
            ms.SetProperty(p);

            Exception t = null;
            try
            {
                MemoryStream out1 = new MemoryStream();
                mps.Write(out1);
                out1.Close();
                byte[] bytes = out1.ToArray();

                PropertySet psr = new PropertySet(bytes);
                Assert.IsTrue(psr.IsSummaryInformation);
                Section sr = (Section)psr.Sections[0];
                String title = (String)sr.GetProperty(PropertyIDMap.PID_TITLE);
                Assert.AreEqual(TITLE, title);
            }
            catch (WritingNotSupportedException e)
            {
                t = e;
            }
            catch (IOException e)
            {
                t = e;
            }
            catch (NoPropertySetStreamException e)
            {
                t = e;
            }
            if (t != null)
                Assert.Fail(t.Message);
        }



        /**
         * Writes a property and Reads it back in.
         *
         * @param variantType The property's variant type.
         * @param value The property's value.
         * @param codepage The codepage to use for writing and Reading.
         * @throws UnsupportedVariantTypeException if the variant is not supported.
         * @throws IOException if an I/O exception occurs.
         * @throws ReadingNotSupportedException 
         * @throws UnsupportedEncodingException 
         */
        private void check(long variantType, Object value,
                           int codepage)
        {
            MemoryStream out1 = new MemoryStream();
            VariantSupport.Write(out1, variantType, value, codepage);
            out1.Close();
            byte[] b = out1.ToArray();
            Object objRead =
                VariantSupport.Read(b, 0, b.Length + LittleEndianConsts.INT_SIZE,
                                    variantType, codepage);
            if (objRead is byte[])
            {
                byte[] valueB = (byte[])value;
                byte[] readB = (byte[])objRead;
                if (valueB.Length != readB.Length)
                    Assert.Fail("Byte arrays are different length - expected " + valueB.Length +
                          " but found " + readB.Length);
                int diff1 = diff(valueB, readB);
                if (diff1 >= 0)
                    Assert.Fail("Byte arrays are different. First different byte is at " +
                         "index " + diff1 + ".");
            }
            else
                if (value != null && !value.Equals(objRead))
                {
                    Assert.Fail("Expected: \"" + value + "\" but was: \"" + objRead +
                         "\". Codepage: " + codepage +
                         (codepage == -1 ?
                          " (" + Encoding.Default.EncodingName + ")." : "."));
                }
                else
                    Assert.AreEqual(value, objRead);
        }



        /**
         * Compares two byte arrays.
         *
         * @param a The first byte array
         * @param b The second byte array
         * @return The index of the first byte that is different. If the byte arrays
         * are equal, -1 is returned.
         */
        private int diff(byte[] a, byte[] b)
        {
            int min = Math.Min(a.Length, b.Length);
            for (int i = 0; i < min; i++)
                if (a[i] != b[i])
                    return i;
            if (a.Length != b.Length)
                return min;
            return -1;
        }



        /**
         * This Test method does a Write and Read back Test with all POI
         * filesystems in the "data" directory by performing the following
         * actions for each file:
         * 
         * <ul>
         * 
         * <li>Read its property Set streams.</li>
         * 
         * <li>Create a new POI filesystem containing the origin file's
         * property Set streams.</li>
         * 
         * <li>Read the property Set streams from the POI filesystem just
         * Created.</li>
         * 
         * <li>Compare each property Set stream with the corresponding one from
         * the origin file and check whether they are equal.</li>
         *
         * </ul>
         */
        [Test]
        public void TestRecreate()
        {
            string[] files = _samples.GetFiles();
            for (int i = 0; i < files.Length; i++)
            {
                string filename = Path.GetFileName(files[i]);
                if (filename.StartsWith("Test") && TestReadAllFiles.checkExclude(filename))
                {
                    try
                    {
                        TestRecreate(new FileInfo(files[i]));
                    }
                    catch (Exception e)
                    {
                        throw new IOException("While handling file " + files[i], e);
                    }
                }
            }
        }



        /**
         * Performs the check described in {@link #TestReCreate()} for a single
         * POI filesystem.
         *
         * @param f the POI filesystem to check
         */
        private void TestRecreate(FileInfo f)
        {
            Console.WriteLine("Recreating file \"" + f.Name + "\"");
            
            /* Read the POI filesystem's property Set streams: */
            POIFile[] psf1 = Util.ReadPropertySets(f);

            /* Create a new POI filesystem containing the origin file's
             * property Set streams: */
            FileInfo copy = TempFile.CreateTempFile(f.Name, "");
            //File.Create(Path.Combine(TestContext.CurrentContext.TestDirectory,f.Name));
            //copy.deleteOnExit();
            FileStream out1 = copy.OpenWrite();
            POIFSFileSystem poiFs = new POIFSFileSystem();
            for (int i = 0; i < psf1.Length; i++)
            {
                Stream in1 =
                    new ByteArrayInputStream(psf1[i].GetBytes());
                PropertySet psIn = PropertySetFactory.Create(in1);
                MutablePropertySet psOut = new MutablePropertySet(psIn);
                MemoryStream psStream =
                    new MemoryStream();
                psOut.Write(psStream);
                psStream.Close();
                byte[] streamData = psStream.ToArray();
                poiFs.CreateDocument(new ByteArrayInputStream(streamData),
                                     psf1[i].GetName());
                poiFs.WriteFileSystem(out1);
            }
            poiFs.Close();
            out1.Close();

            /* Read the property Set streams from the POI filesystem just
             * Created. */
            POIFile[] psf2 = Util.ReadPropertySets(copy);
            for (int i = 0; i < psf2.Length; i++)
            {
                byte[] bytes1 = psf1[i].GetBytes();
                byte[] bytes2 = psf2[i].GetBytes();
                Stream in1 = new ByteArrayInputStream(bytes1);
                Stream in2 = new ByteArrayInputStream(bytes2);
                PropertySet ps1 = PropertySetFactory.Create(in1);
                PropertySet ps2 = PropertySetFactory.Create(in2);

                /* Compare the property Set stream with the corresponding one
                 * from the origin file and check whether they are equal. */
                
                Assert.AreEqual(ps1, ps2, "Equality for file " + f.Name);
            }
            out1.Close();
        }

        /**
         * Tests writing and Reading back a proper dictionary.
         */
        [Test]
        public void TestDictionary()
        {
            FileInfo copy = TempFile.CreateTempFile("Test-HPSF", "ole2");
            MutablePropertySet ps1 = new MutablePropertySet();
            using (FileStream out1 = copy.OpenWrite())
            {
                /* Write: */
                POIFSFileSystem poiFs = new POIFSFileSystem();
                MutableSection s = (MutableSection)ps1.Sections[0];
                Hashtable m = new Hashtable(3, 1.0f);
                m[1] = "String 1";
                m[2] = "String 2";
                m[3] = "String 3";
                s.Dictionary = (m);
                s.SetFormatID(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID1);
                int codepage = CodePageUtil.CP_UNICODE;
                s.SetProperty(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2, codepage);
                poiFs.CreateDocument(ps1.ToInputStream(), "Test");
                poiFs.WriteFileSystem(out1);
                poiFs.Close();
            }
            /* Read back: */
            POIFile[] psf = Util.ReadPropertySets(copy);
            Assert.AreEqual(1, psf.Length);
            byte[] bytes = psf[0].GetBytes();
            Stream in1 = new ByteArrayInputStream(bytes);
            PropertySet ps2 = PropertySetFactory.Create(in1);

            /* Check if the result is a DocumentSummaryInformation stream, as
                * specified. */
            Assert.IsTrue(ps2.IsDocumentSummaryInformation);

            /* Compare the property Set stream with the corresponding one
                * from the origin file and check whether they are equal. */
            Assert.IsTrue(ps1.Equals(ps2));
            
            copy.Delete();
        }

        /**
         * Tests that when using NPOIFS, we can do an in-place write
         *  without needing to stream in + out the whole kitchen sink
         */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestInPlaceNPOIFSWrite()
        {
            NPOIFSFileSystem fs = null;
            DirectoryEntry root = null;
            DocumentNode sinfDoc = null;
            DocumentNode dinfDoc = null;
            SummaryInformation sinf = null;
            DocumentSummaryInformation dinf = null;

            // We need to work on a File for in-place changes, so create a temp one
            FileInfo copy = TempFile.CreateTempFile("Test-HPSF", "ole2");
            //copy.DeleteOnExit();

            // Copy a test file over to a temp location
            Stream inp = _samples.OpenResourceAsStream("TestShiftJIS.doc");
            FileStream out1 = new FileStream(copy.FullName, FileMode.Create);
            IOUtils.Copy(inp, out1);
            inp.Close();
            out1.Close();

            // Open the copy in Read/write mode
            fs = new NPOIFSFileSystem(copy, false);
            root = fs.Root;

            // Read the properties in there
            sinfDoc = (DocumentNode)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            InputStream sinfStream = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream);
            sinfStream.Close();
            Assert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream);
            dinfStream.Close();
            Assert.AreEqual(131077, dinf.OSVersion);

            // Check they start as we expect
            Assert.AreEqual("Reiichiro Hori", sinf.Author);
            Assert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            Assert.AreEqual("\u7b2c1\u7ae0", sinf.Title);

            Assert.AreEqual("", dinf.Company);
            Assert.AreEqual(null, dinf.Manager);

            // Do an in-place replace via an InputStream
            new NPOIFSDocument(sinfDoc).ReplaceContents(sinf.ToInputStream());
            new NPOIFSDocument(dinfDoc).ReplaceContents(dinf.ToInputStream());


            // Check it didn't Get Changed
            sinfDoc = (DocumentNode)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            InputStream sinfStream2 = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream2);
            sinfStream2.Close();
            Assert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream2 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream2);
            dinfStream2.Close();
            Assert.AreEqual(131077, dinf.OSVersion);


            // Start again!
            fs.Close();
            inp = _samples.OpenResourceAsStream("TestShiftJIS.doc");
            out1 = new FileStream(copy.FullName, FileMode.Open, FileAccess.ReadWrite);
            IOUtils.Copy(inp, out1);
            inp.Close();
            out1.Close();

            fs = new NPOIFSFileSystem(new FileStream(copy.FullName, FileMode.Open, FileAccess.ReadWrite),
                null, false, true);
            root = fs.Root;

            // Read the properties in once more
            sinfDoc = (DocumentNode)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            InputStream sinfStream3 = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream3);
            sinfStream3.Close();
            Assert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream3 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream3);
            dinfStream3.Close();
            Assert.AreEqual(131077, dinf.OSVersion);


            // Have them write themselves in-place with no Changes
            Stream soufStream = new NDocumentOutputStream(sinfDoc);
            sinf.Write(soufStream);
            soufStream.Close();
            Stream doufStream = new NDocumentOutputStream(dinfDoc);
            dinf.Write(doufStream);
            doufStream.Close();

            // And also write to some bytes for Checking
            MemoryStream sinfBytes = new MemoryStream();
            sinf.Write(sinfBytes);
            MemoryStream dinfBytes = new MemoryStream();
            dinf.Write(dinfBytes);


            // Check that the filesystem can give us back the same bytes
            sinfDoc = (DocumentNode)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            InputStream sinfStream4 = new NDocumentInputStream(sinfDoc);
            byte[] sinfData = IOUtils.ToByteArray(sinfStream4);
            sinfStream4.Close();
            InputStream dinfStream4 = new NDocumentInputStream(dinfDoc);
            byte[] dinfData = IOUtils.ToByteArray(dinfStream4);
            dinfStream4.Close();
            Assert.That(sinfBytes.ToArray(), new EqualConstraint(sinfData));
            Assert.That(dinfBytes.ToArray(), new EqualConstraint(dinfData));


            // Read back in as-is
            InputStream sinfStream5 = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream5);
            sinfStream5.Close();
            Assert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream5 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream5);
            dinfStream5.Close();
            Assert.AreEqual(131077, dinf.OSVersion);

            Assert.AreEqual("Reiichiro Hori", sinf.Author);
            Assert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            Assert.AreEqual("\u7b2c1\u7ae0", sinf.Title);

            Assert.AreEqual("", dinf.Company);
            Assert.AreEqual(null, dinf.Manager);


            // Now alter a few of them
            sinf.Author = (/*setter*/"Changed Author");
            sinf.Title = (/*setter*/"Le titre \u00e9tait chang\u00e9");
            dinf.Manager = (/*setter*/"Changed Manager");


            // Save this into the filesystem
            Stream soufStream2 = new NDocumentOutputStream(sinfDoc);
            sinf.Write(soufStream2);
            soufStream2.Close();
            Stream doufStream2 = new NDocumentOutputStream(dinfDoc);
            dinf.Write(doufStream2);
            doufStream2.Close();


            // Read them back in again
            sinfDoc = (DocumentNode)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            InputStream sinfStream6 = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream6);
            sinfStream6.Close();
            Assert.AreEqual(131077, sinf.OSVersion);

            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            InputStream dinfStream6 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream6);
            dinfStream6.Close();
            Assert.AreEqual(131077, dinf.OSVersion);

            Assert.AreEqual("Changed Author", sinf.Author);
            Assert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            Assert.AreEqual("Le titre \u00e9tait chang\u00e9", sinf.Title);

            Assert.AreEqual("", dinf.Company);
            Assert.AreEqual("Changed Manager", dinf.Manager);


            // Close the whole filesystem, and open it once more
            fs.WriteFileSystem();
            fs.Close();

            fs = new NPOIFSFileSystem(new FileStream(copy.FullName, FileMode.Open));
            root = fs.Root;

            // Re-check on load
            sinfDoc = (DocumentNode)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            InputStream sinfStream7 = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream7);
            sinfStream7.Close();
            Assert.AreEqual(131077, sinf.OSVersion);

            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            InputStream dinfStream7 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream7);
            dinfStream7.Close();
            Assert.AreEqual(131077, dinf.OSVersion);

            Assert.AreEqual("Changed Author", sinf.Author);
            Assert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            Assert.AreEqual("Le titre \u00e9tait chang\u00e9", sinf.Title);

            Assert.AreEqual("", dinf.Company);
            Assert.AreEqual("Changed Manager", dinf.Manager);


            // Tidy up
            fs.Close();
            copy.Delete();
        }


        /**
         * Tests writing and Reading back a proper dictionary with an invalid
         * codepage. (HPSF Writes Unicode dictionaries only.)
         */
        [Test]
        public void TestDictionaryWithInvalidCodepage()
        {

            using (FileStream copy = File.Create(@".\Test-HPSF.ole2"))
            {
                /* Write: */
                POIFSFileSystem poiFs = new POIFSFileSystem();
                MutablePropertySet ps1 = new MutablePropertySet();
                MutableSection s = (MutableSection)ps1.Sections[0];
                Hashtable m = new Hashtable(3, 1.0f);
                m[1] = "String 1";
                m[2] = "String 2";
                m[3] = "String 3";
                
                try
                {
                    Assert.Throws<IllegalPropertySetDataException>(() => {
                        s.Dictionary = m;
                        s.SetFormatID(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID1);
                        s.SetProperty(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2, 12345);
                        poiFs.CreateDocument(ps1.ToInputStream(), "Test");
                        poiFs.WriteFileSystem(copy);
                    });
                }
                finally
                {
                    poiFs.Close();
                }
            }
            
            
            
            if (File.Exists(@".\Test-HPSF.ole2"))
            {
                File.Delete(@".\Test-HPSF.ole2");
            }
        }


        /**
         * Handles unexpected exceptions in Testcases.
         *
         * @param ex The exception that has been thrown.
         */
        private void handle(Exception ex)
        {
            Assert.Fail("Caused by:" + ex.InnerException.Message);
        }



        /**
         * Returns the display name of the default character Set.
         *
         * @return the display name of the default character Set.
         */
        private static String GetDefaultCharSetName()
        {
            //String charSetName = System.GetProperty("file.encoding");
            //CharSet charSet = CharSet.forName(charSetName);
            return Encoding.Default.EncodingName;
        }



        /**
         * In order to execute Tests with characters beyond US-ASCII, this
         * method checks whether the application is runing in an environment
         * where the default character Set is 16-bit-capable.
         *
         * @return <c>true</c> if the default character Set is 16-bit-capable,
         * else <c>false</c>.
         */
        private bool hasProperDefaultCharSet()
        {
            //String charSetName = System.GetProperty("file.encoding");
            //CharSet charSet = CharSet.forName(charSetName);

            return true;
        }

    }
}
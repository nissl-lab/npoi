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
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    using NPOI.HPSF;
    using NPOI.HPSF.Wellknown;
    using NPOI.POIFS.EventFileSystem;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    using NUnit.Framework;using NUnit.Framework.Legacy;
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
        private static POIDataSamples _samples = POIDataSamples.GetHPSFInstance();
        private static int CODEPAGE_DEFAULT = -1;

        private static string POI_FS = "TestHPSFWritingFunctionality.doc";

        private static int BYTE_ORDER = 0xfffe;
        private static int FORMAT     = 0x0000;
        private static int OS_VERSION = 0x00020A04;
        private static int[] SECTION_COUNT = {1, 2};
        private static bool[] IS_SUMMARY_INFORMATION = {true, false};
        private static bool[] IS_DOCUMENT_SUMMARY_INFORMATION = {false, true};

        private string IMPROPER_DEFAULT_CHARSET_MESSAGE =
            "Your default character Set is " + GetDefaultCharSetName() +
            ". However, this Testcase must be run in an environment " +
            "with a default character Set supporting at least " +
            "8-bit-characters. You can achieve this by Setting the " +
            "LANG environment variable to a proper value, e.g. " +
            "\"de_DE\".";

        POIFile[] poiFiles;

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
                    ByteArrayOutputStream psStream = new ByteArrayOutputStream();
                    ps.Write(psStream);
                    psStream.Close();
                    byte[] streamData = psStream.ToByteArray();
                    poiFs.CreateDocument(new MemoryStream(streamData),
                                         SummaryInformation.DEFAULT_STREAM_NAME);
                    poiFs.WriteFileSystem(out1);
                    out1.Close();
                    Assert.Fail("Should have thrown a NoFormatIDException.");
                }
                catch (Exception ex)
                {
                    ClassicAssert.IsTrue(ex is NoFormatIDException);
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
                s.FormatID = SectionIDMap.SUMMARY_INFORMATION_ID;

                ByteArrayOutputStream psStream = new ByteArrayOutputStream();
                ps.Write(psStream);
                psStream.Close();
                byte[] streamData = psStream.ToByteArray();
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
            si.FormatID = SectionIDMap.SUMMARY_INFORMATION_ID;
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
            ClassicAssert.IsNotNull(psa[0]);
            ClassicAssert.IsTrue(psa[0].IsSummaryInformation);

            Section s = (Section)(psa[0].Sections[0]);
            Object p1 = s.GetProperty(PropertyIDMap.PID_AUTHOR);
            Object p2 = s.GetProperty(PropertyIDMap.PID_TITLE);
            ClassicAssert.AreEqual(AUTHOR, p1);
            ClassicAssert.AreEqual(TITLE, p2);
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
            s1.FormatID = formatID;
            s1.SetProperty(2, SECTION1);
            ps.AddSection(s1);

            MutableSection s2 = new MutableSection();
            s2.FormatID = formatID;
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
            ClassicAssert.IsNotNull(psa[0]);
            Section s = (Section)(psa[0].Sections[0]);
            ClassicAssert.AreEqual(s.FormatID, formatID);
            Object p = s.GetProperty(2);
            ClassicAssert.AreEqual(SECTION1, p);
            s = (Section)(psa[0].Sections[1]);
            p = s.GetProperty(2);
            ClassicAssert.AreEqual(SECTION2, p);

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


        /**
         * Writes and Reads back various variant types and checks whether the
         * stuff that has been Read back Equals the stuff that was written.
         */
        [Test]
        public void TestVariantTypes()
        {
            int codepage = CODEPAGE_DEFAULT;
            Assume.That(hasProperDefaultCharSet(), IMPROPER_DEFAULT_CHARSET_MESSAGE);

            check(Variant.VT_EMPTY, null, codepage);
            check(Variant.VT_BOOL, true, codepage);
            check(Variant.VT_BOOL, false, codepage);
            check( Variant.VT_CF, new byte[] { 8, 0, 0, 0, 1, 0, 0, 0, 1, 2, 3, 4 }, codepage );
            check(Variant.VT_I4, 27, codepage);
            check(Variant.VT_I8, 28L, codepage);
            check(Variant.VT_R8, 29.0d, codepage);
            check(Variant.VT_I4, -27, codepage);
            check(Variant.VT_I8, -28L, codepage);
            check(Variant.VT_R8, -29.0d, codepage);
            check(Variant.VT_FILETIME, new DateTime(1984, 5, 16, 8, 23, 15), codepage);
            check(Variant.VT_I4, Int32.MaxValue, codepage);
            check(Variant.VT_I4, Int32.MinValue, codepage);
            check(Variant.VT_I8, long.MaxValue, codepage);
            check(Variant.VT_I8, long.MinValue, codepage);
            check(Variant.VT_R8, Double.MaxValue, codepage);
            check(Variant.VT_R8, Double.MinValue, codepage);
            checkString(Variant.VT_LPSTR, "\u00e4\u00f6\u00fc\u00df\u00c4\u00d6\u00dc", codepage);
            checkString(Variant.VT_LPWSTR, "\u00e4\u00f6\u00fc\u00df\u00c4\u00d6\u00dc", codepage);
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
            int[] validCodepages = new int[] { CODEPAGE_DEFAULT, CodePageUtil.CP_UTF8, CodePageUtil.CP_UNICODE, CodePageUtil.CP_WINDOWS_1252 };
            for (int i = 0; i < validCodepages.Length; i++)
            {
                int cp = validCodepages[i];
                if (cp == -1 && !hasProperDefaultCharSet())
                {
                    Console.Error.WriteLine(IMPROPER_DEFAULT_CHARSET_MESSAGE +
                         " This Testcase is skipped for the default codepage.");
                    continue;
                }

                long t = (cp == CodePageUtil.CP_UNICODE) ? Variant.VT_LPWSTR : Variant.VT_LPSTR;
                checkString(t, "\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df", cp);
                if (cp == CodePageUtil.CP_UTF16 || cp == CodePageUtil.CP_UTF8) {
                    check(t, "\u79D1\u5B78", cp);
                }
            }

            int[] invalidCodepages = new int[] { 0, 1, 2, 4711, 815 };
            foreach (int cp in invalidCodepages)
            {
                long type = (cp == CodePageUtil.CP_UNICODE) ? Variant.VT_LPWSTR : Variant.VT_LPSTR;
                try
                {
                    checkString(type, "\u00e4\u00f6\u00fc\u00c4\u00d6\u00dc\u00df", cp);
                    ClassicAssert.Fail("UnsupportedEncodingException for codepage " + cp + " expected.");
                }
                catch (UnsupportedEncodingException ex)
                {
                    /* This is the expected behaviour. */
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
            ms.FormatID = SectionIDMap.SUMMARY_INFORMATION_ID;
            MutableProperty p = new MutableProperty();
            p.ID = PropertyIDMap.PID_TITLE;
            p.Type = Variant.VT_LPSTR;
            p.Value = TITLE;
            ms.SetProperty(p);

            ByteArrayOutputStream out1 = new ByteArrayOutputStream();
            mps.Write(out1);
            out1.Close();
            byte[] bytes = out1.ToByteArray();

            PropertySet psr = new PropertySet(bytes);
            ClassicAssert.IsTrue(psr.IsSummaryInformation);
            Section sr = (Section)psr.Sections[0];
            String title = (String)sr.GetProperty(PropertyIDMap.PID_TITLE);
            ClassicAssert.AreEqual(TITLE, title);
        }

        private void checkString(long variantType, string value, int codepage)
        {
            for (int i=0; i<value.Length; i++)
            {
                check(variantType, value.Substring(0, i), codepage);
            }
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
        private void check(long variantType, Object value, int codepage)
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
                POITestCase.AssertEquals((byte[])value, (byte[])objRead);
            }
            else if (value != null && !value.Equals(objRead)) {
                ClassicAssert.AreEqual(value, objRead);
            }
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
                Dictionary<long, string> m = new();
                m[1] = "String 1";
                m[2] = "String 2";
                m[3] = "String 3";
                s.SetDictionary(m);
                s.FormatID = SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID[0];
                int codepage = CodePageUtil.CP_UNICODE;
                s.SetProperty(PropertyIDMap.PID_CODEPAGE, Variant.VT_I2, codepage);
                poiFs.CreateDocument(ps1.ToInputStream(), "Test");
                poiFs.WriteFileSystem(out1);
                poiFs.Close();
            }
            /* Read back: */
            List<POIFile> psf = Util.ReadPropertySets(copy);
            ClassicAssert.AreEqual(1, psf.Count);
            byte[] bytes = psf[0].GetBytes();
            InputStream in1 = new ByteArrayInputStream(bytes);
            PropertySet ps2 = PropertySetFactory.Create(in1);

            /* Check if the result is a DocumentSummaryInformation stream, as
                * specified. */
            ClassicAssert.IsTrue(ps2.IsDocumentSummaryInformation);

            /* Compare the property Set stream with the corresponding one
                * from the origin file and check whether they are equal. */
            ClassicAssert.IsTrue(ps1.Equals(ps2));
            
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
            ClassicAssert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream);
            dinfStream.Close();
            ClassicAssert.AreEqual(131077, dinf.OSVersion);

            // Check they start as we expect
            ClassicAssert.AreEqual("Reiichiro Hori", sinf.Author);
            ClassicAssert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            ClassicAssert.AreEqual("\u7b2c1\u7ae0", sinf.Title);

            ClassicAssert.AreEqual("", dinf.Company);
            ClassicAssert.AreEqual(null, dinf.Manager);

            // Do an in-place replace via an InputStream
            new NPOIFSDocument(sinfDoc).ReplaceContents(sinf.ToInputStream());
            new NPOIFSDocument(dinfDoc).ReplaceContents(dinf.ToInputStream());


            // Check it didn't Get Changed
            sinfDoc = (DocumentNode)root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            InputStream sinfStream2 = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream2);
            sinfStream2.Close();
            ClassicAssert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream2 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream2);
            dinfStream2.Close();
            ClassicAssert.AreEqual(131077, dinf.OSVersion);


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
            ClassicAssert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream3 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream3);
            dinfStream3.Close();
            ClassicAssert.AreEqual(131077, dinf.OSVersion);


            // Have them write themselves in-place with no Changes
            Stream soufStream = new NDocumentOutputStream(sinfDoc);
            sinf.Write(soufStream);
            soufStream.Close();
            Stream doufStream = new NDocumentOutputStream(dinfDoc);
            dinf.Write(doufStream);
            doufStream.Close();

            // And also write to some bytes for Checking
            ByteArrayOutputStream sinfBytes = new ByteArrayOutputStream();
            sinf.Write(sinfBytes);
            ByteArrayOutputStream dinfBytes = new ByteArrayOutputStream();
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
            Assert.That(sinfBytes.ToByteArray(), new EqualConstraint(sinfData));
            Assert.That(dinfBytes.ToByteArray(), new EqualConstraint(dinfData));


            // Read back in as-is
            InputStream sinfStream5 = new NDocumentInputStream(sinfDoc);
            sinf = (SummaryInformation)PropertySetFactory.Create(sinfStream5);
            sinfStream5.Close();
            ClassicAssert.AreEqual(131077, sinf.OSVersion);

            InputStream dinfStream5 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream5);
            dinfStream5.Close();
            ClassicAssert.AreEqual(131077, dinf.OSVersion);

            ClassicAssert.AreEqual("Reiichiro Hori", sinf.Author);
            ClassicAssert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            ClassicAssert.AreEqual("\u7b2c1\u7ae0", sinf.Title);

            ClassicAssert.AreEqual("", dinf.Company);
            ClassicAssert.AreEqual(null, dinf.Manager);


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
            ClassicAssert.AreEqual(131077, sinf.OSVersion);

            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            InputStream dinfStream6 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream6);
            dinfStream6.Close();
            ClassicAssert.AreEqual(131077, dinf.OSVersion);

            ClassicAssert.AreEqual("Changed Author", sinf.Author);
            ClassicAssert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            ClassicAssert.AreEqual("Le titre \u00e9tait chang\u00e9", sinf.Title);

            ClassicAssert.AreEqual("", dinf.Company);
            ClassicAssert.AreEqual("Changed Manager", dinf.Manager);


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
            ClassicAssert.AreEqual(131077, sinf.OSVersion);

            dinfDoc = (DocumentNode)root.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            InputStream dinfStream7 = new NDocumentInputStream(dinfDoc);
            dinf = (DocumentSummaryInformation)PropertySetFactory.Create(dinfStream7);
            dinfStream7.Close();
            ClassicAssert.AreEqual(131077, dinf.OSVersion);

            ClassicAssert.AreEqual("Changed Author", sinf.Author);
            ClassicAssert.AreEqual("Microsoft Word 9.0", sinf.ApplicationName);
            ClassicAssert.AreEqual("Le titre \u00e9tait chang\u00e9", sinf.Title);

            ClassicAssert.AreEqual("", dinf.Company);
            ClassicAssert.AreEqual("Changed Manager", dinf.Manager);


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
                Dictionary<long, string> m = new();
                m[1] = "String 1";
                m[2] = "String 2";
                m[3] = "String 3";
                
                try
                {
                    Assert.Throws<IllegalPropertySetDataException>(() => {
                        s.SetDictionary(m);
                        s.FormatID = SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID[0];
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
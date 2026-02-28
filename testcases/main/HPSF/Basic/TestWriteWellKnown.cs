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
    using System.Globalization;
    using System.IO;

    using NPOI.HPSF;
    using NPOI.HPSF.Wellknown;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    /**
     * Tests HPSF's high-level writing functionality for the well-known property
     * Set "SummaryInformation" and "DocumentSummaryInformation".
     * 
     * @author Rainer Klute
     *     <a href="mailto:klute@rainer-klute.de">klute@rainer-klute.de</a>
     * @since 2006-02-01
     * @version $Id: TestWriteWellKnown.java 489730 2006-12-22 19:18:16Z bayard $
     */
    [TestFixture]
    public class TestWriteWellKnown
    {
        //static string dataDir = @"..\..\..\TestCases\HPSF\data\";
        private static String POI_FS = "TestWriteWellKnown.doc";

        /**
         * @see TestCase#SetUp()
         */
        [SetUp]
        public void SetUp()
        {
            VariantSupport.IsLogUnsupportedTypes = false;
        }

        static string P_APPLICATION_NAME = "ApplicationName";
        static string P_AUTHOR = "Author";
        static int    P_CHAR_COUNT = 4712;
        static string P_COMMENTS = "Comments";
        static DateTime   P_CREATE_DATE_TIME;
        static long   P_EDIT_TIME = 4713 * 1000 * 10;
        static string P_KEYWORDS = "Keywords";
        static string P_LAST_AUTHOR = "LastAuthor";
        static DateTime   P_LAST_PRINTED;
        static DateTime   P_LAST_SAVE_DATE_TIME;
        static int    P_PAGE_COUNT = 4714;
        static string P_REV_NUMBER = "RevNumber";
        static int    P_SECURITY = 1;
        static string P_SUBJECT = "Subject";
        static string P_TEMPLATE = "Template";
        // FIXME (byte array properties not yet implemented): static final byte[] P_THUMBNAIL = new byte[123];
        static string P_TITLE = "Title";
        static int    P_WORD_COUNT = 4715;

        static int     P_BYTE_COUNT = 4716;
        static string  P_CATEGORY = "Category";
        static string  P_COMPANY = "Company";
        // FIXME (byte array properties not yet implemented): static final byte[]  P_DOCPARTS = new byte[123];
        // FIXME (byte array properties not yet implemented): static final byte[]  P_HEADING_PAIR = new byte[123];
        static int     P_HIDDEN_COUNT = 4717;
        static int     P_LINE_COUNT = 4718;
        static bool P_LINKS_DIRTY = true;
        static string  P_MANAGER = "Manager";
        static int     P_MM_CLIP_COUNT = 4719;
        static int     P_NOTE_COUNT = 4720;
        static int     P_PAR_COUNT = 4721;
        static string  P_PRESENTATION_FORMAT = "PresentationFormat";
        static bool P_SCALE = false;
        static int     P_SLIDE_COUNT = 4722;
        static DateTime    now = new DateTime();

        static int POSITIVE_INTEGER = 2222;
        static long POSITIVE_LONG = 3333;
        static Double POSITIVE_DOUBLE = 4444;
        static int NEGATIVE_INTEGER = 2222;
        static long NEGATIVE_LONG = 3333;
        static Double NEGATIVE_DOUBLE = 4444;

        static int MAX_INTEGER = Int32.MaxValue;
        static int MIN_INTEGER = Int32.MinValue;
        static long MAX_LONG = long.MaxValue;
        static long MIN_LONG = long.MinValue;
        static Double MAX_DOUBLE = Double.MaxValue;
        static Double MIN_DOUBLE = Double.MinValue;

        static TestWriteWellKnown()
        {
            DateTime cal = LocaleUtil.GetLocaleCalendar(2000, 6, 6, 6, 6, 6);
            P_CREATE_DATE_TIME = new DateTime(2006, 2, 1, 7, 36, 0);
            //cal.Set(2001, 7, 7, 7, 7, 7);
            P_LAST_PRINTED = new DateTime(2001, 7, 7, 7, 7, 7);
            //cal.Set(2002, 8, 8, 8, 8, 8);
            P_LAST_SAVE_DATE_TIME = new DateTime(2008, 9, 30, 9, 54, 0);
        }

        /// <summary>
        /// <para>
        /// This test method test the writing of properties in the well-known
        /// property Set streams "SummaryInformation" and
        /// "DocumentSummaryInformation" by performing the following steps:
        /// </para>
        /// <list type="number">
        /// <item><description>
        /// Read a test document <em>doc1</em> into a POI filesystem.
        /// </description></item>
        /// <item><description>
        /// Read the summary information stream and the document summary
        /// information stream from the POI filesystem.
        /// </description></item>
        /// <item><description>
        /// Write all properties supported by HPSF to the summary
        /// information (e.g. author, edit date, application name) and to the
        /// document summary information (e.g. company, manager).
        /// </description></item>
        /// <item><description>
        /// Write the summary information stream and the document summary
        /// information stream to the POI filesystem.
        /// </description></item>
        /// <item><description>
        /// Write the POI filesystem to a (temporary) file <em>doc2</em>
        /// and close the latter.
        /// </description></item>
        /// <item><description>
        /// Open <em>doc2</em> for reading and check summary information
        /// and document summary information. All properties written before must be
        /// found in the property streams of <em>doc2</em> and have the correct
        /// values.
        /// </description></item>
        /// <item><description>
        /// Remove all properties supported by HPSF from the summary
        /// information (e.g. author, edit date, application name) and from the
        /// document summary information (e.g. company, manager).
        /// </description></item>
        /// <item><description>
        /// Write the summary information stream and the document summary
        /// information stream to the POI filesystem.
        /// </description></item>
        /// <item><description>
        /// Write the POI filesystem to a (temporary) file <em>doc3</em>
        /// and close the latter.
        /// </description></item>
        /// <item><description>
        /// Open <em>doc3</em> for reading and check summary information
        /// and document summary information. All properties removed before must not
        /// be found in the property streams of <em>doc3</em>.</description></item>
        /// </list>
        /// </summary>
        [Test]
        [Ignore("should re-implement NPOIFSFileSystem")]
        public void TestWriteWellKnown1()
        {
            POIDataSamples _samples = POIDataSamples.GetHPSFInstance();
        
            FileInfo doc1 = TempFile.CreateTempFile("POI_HPSF_Test1.", ".tmp");
            FileInfo doc2 = TempFile.CreateTempFile("POI_HPSF_Test2.", ".tmp");
            FileInfo doc3 = TempFile.CreateTempFile("POI_HPSF_Test3.", ".tmp");

            FileInputStream fis = new FileInputStream(_samples.GetFile(POI_FS));
            FileStream fos = new FileStream(doc1.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            IOUtils.Copy(fis, fos);
            fos.Close();
            fis.Close();
        
            CustomProperties cps1 = Write1stFile(doc1, doc2);
            CustomProperties cps2 = Write2ndFile(doc2, doc3);
            Write3rdFile(doc3, null);
        
            ClassicAssert.AreEqual(cps1, cps2);
        }
    
        /*
         * Write all properties supported by HPSF to the summary information
         * (e.g. author, edit date, application name) and to the document
         * summary information (e.g. company, manager).
         */
        private static CustomProperties Write1stFile(FileInfo fileIn, FileInfo fileOut)
        {
            /* Read a test document <em>doc1</em> into a POI filesystem. */
            NPOIFSFileSystem poifs = new NPOIFSFileSystem(fileIn, false);

            /*
             * Read the summary information stream and the document summary
             * information stream from the POI filesystem.
             *
             * Please note that the result consists of SummaryInformation and
             * DocumentSummaryInformation instances which are in memory only. To
             * make them permanent they have to be written to a POI filesystem
             * explicitly (overwriting the former contents). Then the POI filesystem
             * should be saved to a file.
             */
            SummaryInformation si = GetSummaryInformation(poifs);
            DocumentSummaryInformation dsi = GetDocumentSummaryInformation(poifs);

            si.ApplicationName = P_APPLICATION_NAME;
            si.Author = P_AUTHOR;
            si.CharCount = P_CHAR_COUNT;
            si.Comments = P_COMMENTS;
            si.CreateDateTime = P_CREATE_DATE_TIME;
            si.EditTime = P_EDIT_TIME;
            si.Keywords = P_KEYWORDS;
            si.LastAuthor = P_LAST_AUTHOR;
            si.LastPrinted = P_LAST_PRINTED;
            si.LastSaveDateTime = P_LAST_SAVE_DATE_TIME;
            si.PageCount = P_PAGE_COUNT;
            si.RevNumber = P_REV_NUMBER;
            si.Security = P_SECURITY;
            si.Subject = P_SUBJECT;
            si.Template = P_TEMPLATE;
            // FIXME (byte array properties not yet implemented): si.Thumbnail = P_THUMBNAIL;
            si.Title = P_TITLE;
            si.WordCount = P_WORD_COUNT;

            dsi.ByteCount = P_BYTE_COUNT;
            dsi.Category = P_CATEGORY;
            dsi.Company = P_COMPANY;
            // FIXME (byte array properties not yet implemented): dsi.Docparts = P_DOCPARTS;
            // FIXME (byte array properties not yet implemented): dsi.HeadingPair = P_HEADING_PAIR;
            dsi.HiddenCount = P_HIDDEN_COUNT;
            dsi.LineCount = P_LINE_COUNT;
            dsi.LinksDirty = P_LINKS_DIRTY;
            dsi.Manager = P_MANAGER;
            dsi.MMClipCount = P_MM_CLIP_COUNT;
            dsi.NoteCount = P_NOTE_COUNT;
            dsi.ParCount = P_PAR_COUNT;
            dsi.PresentationFormat = P_PRESENTATION_FORMAT;
            dsi.Scale = P_SCALE;
            dsi.SlideCount = P_SLIDE_COUNT;

            CustomProperties cps = dsi.CustomProperties;
            ClassicAssert.IsNull(cps);
            cps = new CustomProperties();
            cps.Put("Schl\u00fcssel \u00e4",    "Wert \u00e4");
            cps.Put("Schl\u00fcssel \u00e4\u00f6",   "Wert \u00e4\u00f6");
            cps.Put("Schl\u00fcssel \u00e4\u00f6\u00fc",  "Wert \u00e4\u00f6\u00fc");
            cps.Put("Schl\u00fcssel \u00e4\u00f6\u00fc\u00d6", "Wert \u00e4\u00f6\u00fc\u00d6");
            cps.Put("positive_int", POSITIVE_INTEGER);
            cps.Put("positive_Long", POSITIVE_LONG);
            cps.Put("positive_Double", POSITIVE_DOUBLE);
            cps.Put("negative_int", NEGATIVE_INTEGER);
            cps.Put("negative_Long", NEGATIVE_LONG);
            cps.Put("negative_Double", NEGATIVE_DOUBLE);
            cps.Put("Boolean", true);
            cps.Put("Date", now);
            cps.Put("max_int", MAX_INTEGER);
            cps.Put("min_int", MIN_INTEGER);
            cps.Put("max_Long", MAX_LONG);
            cps.Put("min_Long", MIN_LONG);
            cps.Put("max_Double", MAX_DOUBLE);
            cps.Put("min_Double", MIN_DOUBLE);
        
            // Check the keys went in
            ClassicAssert.IsTrue(cps.ContainsKey("Schl\u00fcssel \u00e4"));
            ClassicAssert.IsTrue(cps.ContainsKey("Boolean"));
        
            // Check the values went in
            ClassicAssert.AreEqual("Wert \u00e4", cps.Get("Schl\u00fcssel \u00e4"));
            ClassicAssert.AreEqual(true, cps.Get("Boolean"));
            ClassicAssert.IsTrue(cps.ContainsValue(true));
            ClassicAssert.IsTrue(cps.ContainsValue("Wert \u00e4"));
        
            // Check that things that aren't in aren't in
            ClassicAssert.IsFalse(cps.ContainsKey("False Boolean"));
            ClassicAssert.IsFalse(cps.ContainsValue(false));

            // Save as our custom properties
            dsi.CustomProperties = cps;

        
            /* Write the summary information stream and the document summary
             * information stream to the POI filesystem. */
            si.Write(poifs.Root, SummaryInformation.DEFAULT_STREAM_NAME);
            dsi.Write(poifs.Root, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            /* Write the POI filesystem to a (temporary) file <em>doc2</em>
             * and close the latter. */
            Stream out1 = new FileStream(fileOut.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            poifs.WriteFileSystem(out1);
            out1.Close();
            poifs.Close();
        
            return cps;
        }
    
        /*
         * Open <em>doc2</em> for reading and check summary information and
         * document summary information. All properties written before must be
         * found in the property streams of <em>doc2</em> and have the correct
         * values.
         */
        private static CustomProperties Write2ndFile(FileInfo fileIn, FileInfo fileOut)
        {
            NPOIFSFileSystem poifs = new NPOIFSFileSystem(fileIn, false);
            SummaryInformation si = GetSummaryInformation(poifs);
            DocumentSummaryInformation dsi = GetDocumentSummaryInformation(poifs);

            ClassicAssert.AreEqual(P_APPLICATION_NAME, si.ApplicationName);
            ClassicAssert.AreEqual(P_AUTHOR, si.Author);
            ClassicAssert.AreEqual(P_CHAR_COUNT, si.CharCount);
            ClassicAssert.AreEqual(P_COMMENTS, si.Comments);
            ClassicAssert.AreEqual(P_CREATE_DATE_TIME, si.CreateDateTime);
            ClassicAssert.AreEqual(P_EDIT_TIME, si.EditTime);
            ClassicAssert.AreEqual(P_KEYWORDS, si.Keywords);
            ClassicAssert.AreEqual(P_LAST_AUTHOR, si.LastAuthor);
            ClassicAssert.AreEqual(P_LAST_PRINTED, si.LastPrinted);
            ClassicAssert.AreEqual(P_LAST_SAVE_DATE_TIME, si.LastSaveDateTime);
            ClassicAssert.AreEqual(P_PAGE_COUNT, si.PageCount);
            ClassicAssert.AreEqual(P_REV_NUMBER, si.RevNumber);
            ClassicAssert.AreEqual(P_SECURITY, si.Security);
            ClassicAssert.AreEqual(P_SUBJECT, si.Subject);
            ClassicAssert.AreEqual(P_TEMPLATE, si.Template);
            // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(P_THUMBNAIL, si.Thumbnail);
            ClassicAssert.AreEqual(P_TITLE, si.Title);
            ClassicAssert.AreEqual(P_WORD_COUNT, si.WordCount);

            ClassicAssert.AreEqual(P_BYTE_COUNT, dsi.ByteCount);
            ClassicAssert.AreEqual(P_CATEGORY, dsi.Category);
            ClassicAssert.AreEqual(P_COMPANY, dsi.Company);
            // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(P_, dsi.Docparts);
            // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(P_, dsi.HeadingPair);
            ClassicAssert.AreEqual(P_HIDDEN_COUNT, dsi.HiddenCount);
            ClassicAssert.AreEqual(P_LINE_COUNT, dsi.LineCount);
            ClassicAssert.AreEqual(P_LINKS_DIRTY, dsi.LinksDirty);
            ClassicAssert.AreEqual(P_MANAGER, dsi.Manager);
            ClassicAssert.AreEqual(P_MM_CLIP_COUNT, dsi.MMClipCount);
            ClassicAssert.AreEqual(P_NOTE_COUNT, dsi.NoteCount);
            ClassicAssert.AreEqual(P_PAR_COUNT, dsi.ParCount);
            ClassicAssert.AreEqual(P_PRESENTATION_FORMAT, dsi.PresentationFormat);
            ClassicAssert.AreEqual(P_SCALE, dsi.Scale);
            ClassicAssert.AreEqual(P_SLIDE_COUNT, dsi.SlideCount);

            CustomProperties cps = dsi.CustomProperties;
            ClassicAssert.IsNotNull(cps);
            ClassicAssert.IsNull(cps.Get("No value available"));
            ClassicAssert.AreEqual("Wert \u00e4", cps.Get("Schl\u00fcssel \u00e4"));
            ClassicAssert.AreEqual("Wert \u00e4\u00f6", cps.Get("Schl\u00fcssel \u00e4\u00f6"));
            ClassicAssert.AreEqual("Wert \u00e4\u00f6\u00fc", cps.Get("Schl\u00fcssel \u00e4\u00f6\u00fc"));
            ClassicAssert.AreEqual("Wert \u00e4\u00f6\u00fc\u00d6", cps.Get("Schl\u00fcssel \u00e4\u00f6\u00fc\u00d6"));
            ClassicAssert.AreEqual(POSITIVE_INTEGER, cps.Get("positive_int"));
            ClassicAssert.AreEqual(POSITIVE_LONG, cps.Get("positive_Long"));
            ClassicAssert.AreEqual(POSITIVE_DOUBLE, cps.Get("positive_Double"));
            ClassicAssert.AreEqual(NEGATIVE_INTEGER, cps.Get("negative_int"));
            ClassicAssert.AreEqual(NEGATIVE_LONG, cps.Get("negative_Long"));
            ClassicAssert.AreEqual(NEGATIVE_DOUBLE, cps.Get("negative_Double"));
            ClassicAssert.AreEqual(true, cps.Get("Boolean"));
            ClassicAssert.AreEqual(now, cps.Get("Date"));
            ClassicAssert.AreEqual(MAX_INTEGER, cps.Get("max_int"));
            ClassicAssert.AreEqual(MIN_INTEGER, cps.Get("min_int"));
            ClassicAssert.AreEqual(MAX_LONG, cps.Get("max_Long"));
            ClassicAssert.AreEqual(MIN_LONG, cps.Get("min_Long"));
            ClassicAssert.AreEqual(MAX_DOUBLE, cps.Get("max_Double"));
            ClassicAssert.AreEqual(MIN_DOUBLE, cps.Get("min_Double"));

            /* Remove all properties supported by HPSF from the summary
             * information (e.g. author, edit date, application name) and from the
             * document summary information (e.g. company, manager). */
            si.RemoveApplicationName();
            si.RemoveAuthor();
            si.RemoveCharCount();
            si.RemoveComments();
            si.RemoveCreateDateTime();
            si.RemoveEditTime();
            si.RemoveKeywords();
            si.RemoveLastAuthor();
            si.RemoveLastPrinted();
            si.RemoveLastSaveDateTime();
            si.RemovePageCount();
            si.RemoveRevNumber();
            si.RemoveSecurity();
            si.RemoveSubject();
            si.RemoveTemplate();
            si.RemoveThumbnail();
            si.RemoveTitle();
            si.RemoveWordCount();

            dsi.RemoveByteCount();
            dsi.RemoveCategory();
            dsi.RemoveCompany();
            dsi.RemoveCustomProperties();
            dsi.RemoveDocparts();
            dsi.RemoveHeadingPair();
            dsi.RemoveHiddenCount();
            dsi.RemoveLineCount();
            dsi.RemoveLinksDirty();
            dsi.RemoveManager();
            dsi.RemoveMMClipCount();
            dsi.RemoveNoteCount();
            dsi.RemoveParCount();
            dsi.RemovePresentationFormat();
            dsi.RemoveScale();
            dsi.RemoveSlideCount();

            // Write the summary information stream and the document summary
            // information stream to the POI filesystem.
            si.Write(poifs.Root, SummaryInformation.DEFAULT_STREAM_NAME);
            dsi.Write(poifs.Root, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            // Write the POI filesystem to a (temporary) file doc3 and close the latter.
            Stream out1 = new FileStream(fileOut.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            poifs.WriteFileSystem(out1);
            out1.Close();
            poifs.Close();
        
            return cps;
        }
    
        /*
         * Open {@code doc3} for reading and check summary information
         * and document summary information. All properties removed before must not
         * be found in the property streams of {@code doc3}.
         */
        private static CustomProperties Write3rdFile(FileInfo fileIn, FileInfo fileOut)
        {
            NPOIFSFileSystem poifs = new NPOIFSFileSystem(fileIn, false);
            SummaryInformation si = GetSummaryInformation(poifs);
            DocumentSummaryInformation dsi = GetDocumentSummaryInformation(poifs);

            ClassicAssert.IsNull(si.ApplicationName);
            ClassicAssert.IsNull(si.Author);
            ClassicAssert.AreEqual(0, si.CharCount);
            ClassicAssert.IsTrue(si.WasNull);
            ClassicAssert.IsNull(si.Comments);
            ClassicAssert.IsNull(si.CreateDateTime);
            ClassicAssert.AreEqual(0, si.EditTime);
            ClassicAssert.IsTrue(si.WasNull);
            ClassicAssert.IsNull(si.Keywords);
            ClassicAssert.IsNull(si.LastAuthor);
            ClassicAssert.IsNull(si.LastPrinted);
            ClassicAssert.IsNull(si.LastSaveDateTime);
            ClassicAssert.AreEqual(0, si.PageCount);
            ClassicAssert.IsTrue(si.WasNull);
            ClassicAssert.IsNull(si.RevNumber);
            ClassicAssert.AreEqual(0, si.Security);
            ClassicAssert.IsTrue(si.WasNull);
            ClassicAssert.IsNull(si.Subject);
            ClassicAssert.IsNull(si.Template);
            ClassicAssert.IsNull(si.Thumbnail);
            ClassicAssert.IsNull(si.Title);
            ClassicAssert.AreEqual(0, si.WordCount);
            ClassicAssert.IsTrue(si.WasNull);

            ClassicAssert.AreEqual(0, dsi.ByteCount);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.IsNull(dsi.Category);
            ClassicAssert.IsNull(dsi.CustomProperties);
            // FIXME (byte array properties not yet implemented): ClassicAssert.IsNull(dsi.Docparts);
            // FIXME (byte array properties not yet implemented): ClassicAssert.IsNull(dsi.HeadingPair);
            ClassicAssert.AreEqual(0, dsi.HiddenCount);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.AreEqual(0, dsi.LineCount);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.IsFalse(dsi.LinksDirty);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.IsNull(dsi.Manager);
            ClassicAssert.AreEqual(0, dsi.MMClipCount);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.AreEqual(0, dsi.NoteCount);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.AreEqual(0, dsi.ParCount);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.IsNull(dsi.PresentationFormat);
            ClassicAssert.IsFalse(dsi.Scale);
            ClassicAssert.IsTrue(dsi.WasNull);
            ClassicAssert.AreEqual(0, dsi.SlideCount);
            ClassicAssert.IsTrue(dsi.WasNull);
            poifs.Close();
        
            return dsi.CustomProperties;
        }

        internal static SummaryInformation GetSummaryInformation(NPOIFSFileSystem poifs)
        {
            DocumentInputStream dis = poifs.CreateDocumentInputStream(SummaryInformation.DEFAULT_STREAM_NAME);
            PropertySet ps = new PropertySet(dis);
            SummaryInformation si = new SummaryInformation(ps);
            dis.Close();
            return si;
        }
    
        internal static DocumentSummaryInformation GetDocumentSummaryInformation(NPOIFSFileSystem poifs)
        {
            if (!poifs.Root.HasEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME))
            {
                return null;
            }

            DocumentInputStream dis = poifs.CreateDocumentInputStream(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            PropertySet ps = new PropertySet(dis);
            DocumentSummaryInformation dsi = new DocumentSummaryInformation(ps);
            dis.Close();
            return dsi;
        }

        /// <summary>
        /// Tests basic custom property features.
        /// </summary>
        [Test]
        public void TestCustomerProperties()
        {
            string KEY = "Schl\u00fcssel \u00e4";
            string VALUE_1 = "Wert 1";
            string VALUE_2 = "Wert 2";

            CustomProperty cp;
            CustomProperties cps = new CustomProperties();
            ClassicAssert.AreEqual(0, cps.Size);

            /* After adding a custom property the size must be 1 and it must be
             * possible to extract the custom property from the map. */
            cps.Put(KEY, VALUE_1);
            ClassicAssert.AreEqual(1, cps.Size);
            object v1 = cps.Get(KEY);
            ClassicAssert.AreEqual(VALUE_1, v1);

            /* After adding a custom property with the same name the size must still
             * be one. */
            cps.Put(KEY, VALUE_2);
            ClassicAssert.AreEqual(1, cps.Size);
            object v2 = cps.Get(KEY);
            ClassicAssert.AreEqual(VALUE_2, v2);

            /* Removing the custom property must return the remove property and
             * reduce the size to 0. */
            cp = (CustomProperty) cps.Remove(KEY);
            ClassicAssert.AreEqual(KEY, cp.Name);
            ClassicAssert.AreEqual(VALUE_2, cp.Value);
            ClassicAssert.AreEqual(0, cps.Size);
        }



        /// <summary>
        /// Tests reading custom properties from a section including reading
        /// custom properties which are not pure.
        /// </summary>
        [Test]
        public void TestGetCustomerProperties()
        {
            int ID_1 = 2;
            int ID_2 = 3;
            string NAME_1 = "Schl\u00fcssel \u00e4";
            string VALUE_1 = "Wert 1";
            Dictionary<long,String> dictionary = new Dictionary<long, String>();

            DocumentSummaryInformation dsi = PropertySetFactory.NewDocumentSummaryInformation();
            CustomProperties cps;
            MutableSection s;

            /* A document summary information Set stream by default does have custom properties. */
            cps = dsi.CustomProperties;
            ClassicAssert.IsNull(cps);

            /* Test an empty custom properties Set. */
            s = new MutableSection();
            s.FormatID = SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID[1];
            // s.Codepage = CodePageUtil.CP_UNICODE;
            dsi.AddSection(s);
            cps = dsi.CustomProperties;
            ClassicAssert.AreEqual(0, cps.Size);

            /* Add a custom property. */
            MutableProperty p = new MutableProperty();
            p.ID = ID_1;
            p.Type = Variant.VT_LPWSTR;
            p.Value = VALUE_1;
            s.SetProperty(p);
            dictionary[ID_1] = NAME_1;
            s.SetDictionary(dictionary);
            cps = dsi.CustomProperties;
            ClassicAssert.AreEqual(1, cps.Size);
            ClassicAssert.IsTrue(cps.IsPure);

            /* Add another custom property. */
            s.SetProperty(ID_2, Variant.VT_LPWSTR, VALUE_1);
            dictionary[ID_2] = NAME_1;
            s.SetDictionary(dictionary);
            cps = dsi.CustomProperties;
            ClassicAssert.AreEqual(1, cps.Size);
            ClassicAssert.IsFalse(cps.IsPure);
        }
    }
}
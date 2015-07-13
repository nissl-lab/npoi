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

namespace NPOI.XWPF.Extractor
{
    using System;



    using NPOI.XWPF;
    using NPOI.XWPF.UserModel;
    using NPOI.XWPF.Extractor;
    using NUnit.Framework;
    using System.Diagnostics;
    using System.Text.RegularExpressions;

    /**
     * Tests for HXFWordExtractor
     */
    [TestFixture]
    public class TestXWPFWordExtractor
    {

        /**
         * Get text out of the simple file
         * @throws IOException 
         */
        [Test]
        public void TestGetSimpleText()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);

            // Check contents
            Assert.IsTrue(text.StartsWith(
                    "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Nunc at risus vel erat tempus posuere. Aenean non ante. Suspendisse vehicula dolor sit amet odio."
            ));
            Assert.IsTrue(text.EndsWith(
                    "Phasellus ultricies mi nec leo. Sed tempus. In sit amet lorem at velit faucibus vestibulum.\n"
            ));

            // Check number of paragraphs
            int ps = 0;
            char[] t = text.ToCharArray();
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] == '\n')
                {
                    ps++;
                }
            }
            Assert.AreEqual(3, ps);
        }

        /**
         * Tests Getting the text out of a complex file
         * @throws IOException 
         */
        [Test]
        public void TestGetComplexText()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("IllustrativeCases.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);

            char euro = '\u20ac';
            Debug.WriteLine("'" + text.Substring(text.Length - 40) + "'");

            //Check contents
            Assert.IsTrue(text.StartsWith(
                    "  \n(V) ILLUSTRATIVE CASES\n\n"
            ));
            Assert.IsTrue(text.Contains(
                    "As well as gaining " + euro + "90 from child benefit increases, he will also receive the early childhood supplement of " + euro + "250 per quarter for Vincent for the full four quarters of the year.\n\n\n\n"// \n\n\n"
            ));
            Assert.IsTrue(text.EndsWith(
                    "11.4%\t\t90\t\t\t\t\t250\t\t1,310\t\n\n \n\n\n"
            ));

            // Check number of paragraphs
            int ps = 0;
            char[] t = text.ToCharArray();
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] == '\n')
                {
                    ps++;
                }
            }
            Assert.AreEqual(134, ps);
        }

        [Test]
        public void TestGetWithHyperlinks()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("TestDocument.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            // Now check contents
            extractor.SetFetchHyperlinks(false);
            Assert.AreEqual(
                    "This is a test document.\nThis bit is in bold and italic\n" +
                    "Back to normal\n" +
                    "This contains BOLD, ITALIC and BOTH, as well as RED and YELLOW text.\n" +
                    "We have a hyperlink here, and another.\n",
                    extractor.Text
            );

            // One hyperlink is a real one, one is just to the top of page
            extractor.SetFetchHyperlinks(true);
            Assert.AreEqual(
                    "This is a test document.\nThis bit is in bold and italic\n" +
                    "Back to normal\n" +
                    "This contains BOLD, ITALIC and BOTH, as well as RED and YELLOW text.\n" +
                    "We have a hyperlink <http://poi.apache.org/> here, and another.\n",
                    extractor.Text
            );
        }

        [Test]
        public void TestHeadersFooters()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("ThreeColHeadFoot.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            Assert.AreEqual(
                    "First header column!\tMid header\tRight header!\n" +
                            "This is a sample word document. It has two pages. It has a three column heading, and a three column footer\n" +
                            "\n" +
                            "HEADING TEXT\n" +
                            "\n" +
                            "More on page one\n" +
                            "\n\n" +
                            "End of page 1\n\n\n" +
                            "This is page two. It also has a three column heading, and a three column footer.\n" +
                            "Footer Left\tFooter Middle\tFooter Right\n",
                    extractor.Text
            );

            // Now another file, expect multiple headers
            //  and multiple footers
            doc = XWPFTestDataSamples.OpenSampleDocument("DiffFirstPageHeadFoot.docx");
            extractor = new XWPFWordExtractor(doc);
            extractor =
                    new XWPFWordExtractor(doc);
            //extractor.Text;

            Assert.AreEqual(
                    "I am the header on the first page, and I" + '\u2019' + "m nice and simple\n" +
                            "First header column!\tMid header\tRight header!\n" +
                            "This is a sample word document. It has two pages. It has a simple header and footer, which is different to all the other pages.\n" +
                            "\n" +
                            "HEADING TEXT\n" +
                            "\n" +
                            "More on page one\n" +
                            "\n\n" +
                            "End of page 1\n\n\n" +
                            "This is page two. It also has a three column heading, and a three column footer.\n" +
                            "The footer of the first page\n" +
                            "Footer Left\tFooter Middle\tFooter Right\n",
                    extractor.Text
            );
        }

        [Test]
        public void TestFootnotes()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("footnotes.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
            String text = extractor.Text;
            Assert.IsTrue(extractor.Text.Contains("snoska"));
            Assert.IsTrue(text.Contains("Eto ochen prostoy[footnoteRef:1] text so snoskoy"));
        }


        [Test]
        public void TestTableFootnotes()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("table_footnotes.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            Assert.IsTrue(extractor.Text.Contains("snoska"));
        }

        [Test]
        public void TestFormFootnotes()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("form_footnotes.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            String text = extractor.Text;
            Assert.IsTrue(text.Contains("testdoc"), "Unable to find expected word in text\n" + text);
            Assert.IsTrue(text.Contains("test phrase"), "Unable to find expected word in text\n" + text);
        }

        [Test]
        public void TestEndnotes()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("endnotes.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
            string text = extractor.Text;
            Assert.IsTrue(text.Contains("XXX"));
            Assert.IsTrue(text.Contains("tilaka [endnoteRef:2]or 'tika'"));
        }

        [Test]
        public void TestInsertedDeletedText()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("delins.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            Assert.IsTrue(extractor.Text.Contains("pendant worn"));
            Assert.IsTrue(extractor.Text.Contains("extremely well"));
        }

        [Test]
        public void TestParagraphHeader()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Headers.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            Assert.IsTrue(extractor.Text.Contains("Section 1"));
            Assert.IsTrue(extractor.Text.Contains("Section 2"));
            Assert.IsTrue(extractor.Text.Contains("Section 3"));
        }

        /**
         * Test that we can open and process .docm
         *  (macro enabled) docx files (bug #45690)
         * @throws IOException 
         */
        [Test]
        public void TestDOCMFiles()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("45690.docm");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            Assert.IsTrue(extractor.Text.Contains("2004"));
            Assert.IsTrue(extractor.Text.Contains("2008"));
            Assert.IsTrue(extractor.Text.Contains("(120 "));
        }

        /**
         * Test that we handle things like tabs and
         *  carriage returns properly in the text that
         *  we're extracting (bug #49189)
         * @throws IOException 
         */
        [Test]
        public void TestDocTabs()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("WithTabs.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            // Check bits
            Assert.IsTrue(extractor.Text.Contains("a"));
            Assert.IsTrue(extractor.Text.Contains("\t"));
            Assert.IsTrue(extractor.Text.Contains("b"));

            // Now check the first paragraph in total
            Assert.IsTrue(extractor.Text.Contains("a\tb\n"));
        }

        /**
         * The output should not contain field codes, e.g. those specified in the
         * w:instrText tag (spec sec. 17.16.23)
         * @throws IOException 
         */
        [Test]
        public void TestNoFieldCodes()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("FieldCodes.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);
            Assert.IsFalse(text.Contains("AUTHOR"));
            Assert.IsFalse(text.Contains("CREATEDATE"));
        }

        /**
         * The output should contain the values of simple fields, those specified
         * with the fldSimple element (spec sec. 17.16.19)
         * @throws IOException 
         */
        [Test]
        public void TestFldSimpleContent()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("FldSimple.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);
            Assert.IsTrue(text.Contains("FldSimple.docx"));
        }

        /**
         * Test for parsing document with Drawings to prevent
         * NoClassDefFoundError for CTAnchor in XWPFRun
         */
        [Test]
        public void TestDrawings()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("drawing.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);
        }

        /**
         * Test for basic extraction of SDT content
         * @throws IOException
         */
        [Test]
        public void TestSimpleControlContent()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Bug54849.docx");
            String[] targs = new String[]{
                "header_rich_text",
                "rich_text",
                "rich_text_pre_table\nrich_text_cell1\t\t\t\n\t\t\t\n\t\t\t\n\nrich_text_post_table",
                "plain_text_no_newlines",
                "plain_text_with_newlines1\nplain_text_with_newlines2\n",
                "watermelon\n",
                "dirt\n",
                "4/16/2013\n",
                "rich_text_in_cell",
                "abc",
                "rich_text_in_paragraph_in_cell",
                "footer_rich_text",
                "footnote_sdt",
                "endnote_sdt"
        };
            XWPFWordExtractor ex = new XWPFWordExtractor(doc);
            String s = ex.Text.ToLower();
            int hits = 0;

            foreach (String targ in targs)
            {
                bool hitted = false;
                if (s.IndexOf(targ) > -1)
                {
                    hitted = true;
                    hits++;
                }
                Assert.AreEqual(true, hitted, "controlled content loading-" + targ);
            }
            Assert.AreEqual(targs.Length, hits, "controlled content loading hit count");

            ex.Close();

            doc = XWPFTestDataSamples.OpenSampleDocument("Bug54771a.docx");
            targs = new String[]{
                "bb",
                "test subtitle\n",
                "test user\n",
        };
            ex = new XWPFWordExtractor(doc);
            s = ex.Text.ToLower();

            //At one point in development there were three copies of the text.
            //This ensures that there is only one copy.
            MatchCollection mc;
            int hit;
            foreach (String targ in targs)
            {
                mc = Regex.Matches(s, targ);
                hit = 0;
                foreach (Match m in mc)
                {
                    if (m.Success)
                        hit++;
                }
                Assert.AreEqual(1, hit, "controlled content loading-" + targ);
            }
            //"test\n" appears twice: once as the "title" and once in the text.
            //This also happens when you save this document as text from MSWord.
            mc = Regex.Matches(s, "test\n");
            hit = 0;
            foreach (Match m in mc)
            {
                if (m.Success)
                    hit++;
            }
            Assert.AreEqual(2, hit, "test<N>");
            ex.Close();

        }

        /** No Header or Footer in document */
        [Test]
        public void TestBug55733()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("55733.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            // Check it gives text without error
            string text = extractor.Text;
            extractor.Close();
        }

        [Test]
        public void TestCheckboxes()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("checkboxes.docx");
            Console.WriteLine(doc);
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            Assert.AreEqual("This is a small test for checkboxes \nunchecked: |_| \n" +
                         "Or checked: |X|\n\n\n\n\n" +
                         "Test a checkbox within a textbox: |_| -> |X|\n\n\n" +
                         "In Table:\n|_|\t|X|\n\n\n" +
                         "In Sequence:\n|X||_||X|\n", extractor.Text);
            extractor.Close();
        }

    }
}

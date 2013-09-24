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

using NPOI.HWPF.Extractor;

using TestCases.HWPF;
using System;
using System.Text;
using TestCases;
using NPOI.POIFS.FileSystem;
using NPOI.HWPF;
using System.Text.RegularExpressions;
using NUnit.Framework;
namespace TestCases.HWPF.Extractor
{
    /**
     * Test the different routes to extracting text
     *
     * @author Nick Burch (nick at torchbox dot com)
     */
    [TestFixture]
    public class TestWordExtractor
    {
        private String[] p_text1 = new String[] {
			"This is a simple word document\r\n",
			"\r\n",
			"It has a number of paragraphs in it\r\n",
			"\r\n",
			"Some of them even feature bold, italic and underlined text\r\n",
			"\r\n",
			"\r\n",
			"This bit is in a different font and size\r\n",
			"\r\n",
			"\r\n",
			"This bit features some red text.\r\n",
			"\r\n",
			"\r\n",
			"It is otherwise very very boring.\r\n"
	};
        private String p_text1_block = "";

        // Well behaved document
        private WordExtractor extractor;
        // Slightly iffy document
        private WordExtractor extractor2;
        // A word doc embeded in an excel file
        private String filename3;

        // With header and footer
        private String filename4;
        // With unicode header and footer
        private String filename5;
        // With footnote
        private String filename6;

        [SetUp]
        public void SetUp()
        {

            String filename = "test2.doc";
            String filename2 = "test.doc";
            filename3 = "excel_with_embeded.xls";
            filename4 = "ThreeColHeadFoot.doc";
            filename5 = "HeaderFooterUnicode.doc";
            filename6 = "footnote.doc";
            POIDataSamples docTests = POIDataSamples.GetDocumentInstance();
            extractor = new WordExtractor(docTests.OpenResourceAsStream(filename));
            extractor2 = new WordExtractor(docTests.OpenResourceAsStream(filename2));

            // Build splat'd out text version
            for (int i = 0; i < p_text1.Length; i++)
            {
                p_text1_block += p_text1[i];
            }
        }

        /**
         * Test paragraph based extraction
         */
        [Test]
        public void TestExtractFromParagraphs()
        {
            String[] text = extractor.ParagraphText;

            Assert.AreEqual(p_text1.Length, text.Length);
            for (int i = 0; i < p_text1.Length; i++)
            {
                Assert.AreEqual(p_text1[i], text[i]);
            }

            // Lots of paragraphs with only a few lines in them
            Assert.AreEqual(24, extractor2.ParagraphText.Length);
            Assert.AreEqual("as d\r\n", extractor2.ParagraphText[16]);
            Assert.AreEqual("as d\r\n", extractor2.ParagraphText[17]);
            Assert.AreEqual("as d\r\n", extractor2.ParagraphText[18]);
        }

        /**
         * Test the paragraph -> flat extraction
         */
        [Test]
        public void TestText()
        {
            Assert.AreEqual(p_text1_block, extractor.Text);
            Regex regex = new Regex("[\\r\\n]");
            
            
            // For the 2nd, should give similar answers for
            //  the two methods, differing only in line endings
            Assert.AreEqual(
                  regex.Replace(extractor2.TextFromPieces, ""),
                  regex.Replace(extractor2.Text, ""));
        }

        /**
         * Test textPieces based extraction
         */
        [Test]
        public void TestExtractFromTextPieces()
        {
            String text = extractor.TextFromPieces;
            Assert.AreEqual(p_text1_block, text);
        }


        /**
         * Test that we can get data from two different
         *  embeded word documents
         * @throws Exception
         */
        [Test]
        public void TestExtractFromEmbeded()
        {
            POIFSFileSystem fs = new POIFSFileSystem(POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream(filename3));
            HWPFDocument doc;
            WordExtractor extractor3;

            DirectoryNode dirA = (DirectoryNode)fs.Root.GetEntry("MBD0000A3B7");
            DirectoryNode dirB = (DirectoryNode)fs.Root.GetEntry("MBD0000A3B2");

            // Should have WordDocument and 1Table
            Assert.IsNotNull(dirA.GetEntry("1Table"));
            Assert.IsNotNull(dirA.GetEntry("WordDocument"));

            Assert.IsNotNull(dirB.GetEntry("1Table"));
            Assert.IsNotNull(dirB.GetEntry("WordDocument"));

            // Check each in turn
            doc = new HWPFDocument(dirA, fs);
            extractor3 = new WordExtractor(doc);

            Assert.IsNotNull(extractor3.Text);
            Assert.IsTrue(extractor3.Text.Length > 20);
            Assert.AreEqual("I am a sample document\r\nNot much on me\r\nI am document 1\r\n", extractor3
                    .Text);
            Assert.AreEqual("Sample Doc 1", extractor3.SummaryInformation.Title);
            Assert.AreEqual("Sample Test", extractor3.SummaryInformation.Subject);

            doc = new HWPFDocument(dirB, fs);
            extractor3 = new WordExtractor(doc);

            Assert.IsNotNull(extractor3.Text);
            Assert.IsTrue(extractor3.Text.Length > 20);
            Assert.AreEqual("I am another sample document\r\nNot much on me\r\nI am document 2\r\n",
                    extractor3.Text);
            Assert.AreEqual("Sample Doc 2", extractor3.SummaryInformation.Title);
            Assert.AreEqual("Another Sample Test", extractor3.SummaryInformation.Subject);
        }
        [Test]
        public void TestWithHeader()
        {
            // Non-unicode
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile(filename4);
            extractor = new WordExtractor(doc);

            Assert.AreEqual("First header column!\tMid header Right header!\n", extractor.HeaderText);

            String text = extractor.Text;
            Assert.IsTrue(text.IndexOf("First header column!") > -1);

            // Unicode
            doc = HWPFTestDataSamples.OpenSampleFile(filename5);
            extractor = new WordExtractor(doc);

            Assert.AreEqual("This is a simple header, with a \u20ac euro symbol in it.\n\n", extractor
                    .HeaderText);
            text = extractor.Text;
            Assert.IsTrue(text.IndexOf("This is a simple header") > -1);
        }
        [Test]
        public void TestWithFooter()
        {
            // Non-unicode
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile(filename4);
            extractor = new WordExtractor(doc);

            Assert.AreEqual("Footer Left\tFooter Middle Footer Right\n", extractor.FooterText);

            String text = extractor.Text;
            Assert.IsTrue(text.IndexOf("Footer Left") > -1);

            // Unicode
            doc = HWPFTestDataSamples.OpenSampleFile(filename5);
            extractor = new WordExtractor(doc);

            Assert.AreEqual("The footer, with Moli\u00e8re, has Unicode in it.\n", extractor
                    .FooterText);
            text = extractor.Text;
            Assert.IsTrue(text.IndexOf("The footer, with") > -1);
        }
        [Test]
        public void TestFootnote()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile(filename6);
            extractor = new WordExtractor(doc);

            String[] text = extractor.FootnoteText;
            StringBuilder b = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                b.Append(text[i]);
            }

            Assert.IsTrue(b.ToString().Contains("TestFootnote"));
        }
        [Test]
        public void TestEndnote()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile(filename6);
            extractor = new WordExtractor(doc);

            String[] text = extractor.EndnoteText;
            StringBuilder b = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                b.Append(text[i]);
            }

            Assert.IsTrue(b.ToString().Contains("TestEndnote"));
        }
        [Test]
        public void TestComments()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile(filename6);
            extractor = new WordExtractor(doc);

            String[] text = extractor.CommentsText;
            StringBuilder b = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                b.Append(text[i]);
            }

            Assert.IsTrue(b.ToString().Contains("TestComment"));
        }
        [Test]
        public void TestWord95()
        {
            // Too old for the default
            try
            {
                extractor = new WordExtractor(
                        POIDataSamples.GetDocumentInstance().OpenResourceAsStream("Word95.doc")
                );
                Assert.Fail();
            }
            catch (OldWordFileFormatException ) { }

            // Can work with the special one
            Word6Extractor w6e = new Word6Extractor(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("Word95.doc")
            );
            String text = w6e.Text;

            Assert.IsTrue(text.Contains("The quick brown fox jumps over the lazy dog"));
            Assert.IsTrue(text.Contains("Paragraph 2"));
            Assert.IsTrue(text.Contains("Paragraph 3. Has some RED text and some BLUE BOLD text in it"));
            Assert.IsTrue(text.Contains("Last (4th) paragraph"));

            String[] tp = w6e.ParagraphText;
            Assert.AreEqual(7, tp.Length);
            Assert.AreEqual("The quick brown fox jumps over the lazy dog\r\n", tp[0]);
            Assert.AreEqual("\r\n", tp[1]);
            Assert.AreEqual("Paragraph 2\r\n", tp[2]);
            Assert.AreEqual("\r\n", tp[3]);
            Assert.AreEqual("Paragraph 3. Has some RED text and some BLUE BOLD text in it.\r\n", tp[4]);
            Assert.AreEqual("\r\n", tp[5]);
            Assert.AreEqual("Last (4th) paragraph.\r\n", tp[6]);
        }
        [Test]
        public void TestWord6()
        {
            // Too old for the default
            try
            {
                extractor = new WordExtractor(
                        POIDataSamples.GetDocumentInstance().OpenResourceAsStream("Word6.doc")
                );
                Assert.Fail();
            }
            catch (OldWordFileFormatException) { }

            Word6Extractor w6e = new Word6Extractor(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("Word6.doc")
            );
            String text = w6e.Text;

            Assert.IsTrue(text.Contains("The quick brown fox jumps over the lazy dog"));

            String[] tp = w6e.ParagraphText;
            Assert.AreEqual(1, tp.Length);
            Assert.AreEqual("The quick brown fox jumps over the lazy dog\r\n", tp[0]);
        }
        [Test]
        public void TestFastSaved()
        {
            extractor = new WordExtractor(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("rasp.doc")
            );

            String text = extractor.Text;
            Assert.IsTrue(text.Contains("\u0425\u0425\u0425\u0425\u0425"));
            Assert.IsTrue(text.Contains("\u0423\u0423\u0423\u0423\u0423"));
        }
        [Test]
        public void TestFirstParagraphFix()
        {
            extractor = new WordExtractor(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("Bug48075.doc")
            );

            String text = extractor.Text;

            Assert.IsTrue(text.StartsWith("\u041f\u0440\u0438\u043b\u043e\u0436\u0435\u043d\u0438\u0435"));
        }
    }

}
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

using NPOI.HWPF.UserModel;

using System;
using System.Collections.Generic;
using NPOI.HWPF.Model;
using NPOI.HWPF;
using NUnit.Framework;
namespace TestCases.HWPF.UserModel
{

    /**
     * Tests to ensure that our ranges end up with
     *  the right text in them, and the right font/styling
     *  properties applied to them.
     */
    [TestFixture]
    public class TestRangeProperties
    {
        private static char page_break = (char)12;

        private static String u_page_1 =
            "This is a fairly simple word document, over two pages, with headers and footers.\r" +
            "The trick with this one is that it contains some Unicode based strings in it.\r" +
            "Firstly, some currency symbols:\r" +
            "\tGBP - \u00a3\r" +
            "\tEUR - \u20ac\r" +
            "Now, we\u2019ll have some French text, in bold and big:\r" +
            "\tMoli\u00e8re\r" +
            "And some normal French text:\r" +
            "\tL'Avare ou l'\u00c9cole du mensonge\r" +
            "That\u2019s it for page one\r"
        ;
        private static String u_page_2 =
            "This is page two. Les Pr\u00e9cieuses ridicules. The end.\r"
        ;

        private static String a_page_1 =
            "I am a test document\r" +
            "This is page 1\r" +
            "I am Calibri (Body) in font size 11\r"
        ;
        private static String a_page_2 =
            "This is page two\r" +
            "It\u2019s Arial Black in 16 point\r" +
            "It\u2019s also in blue\r"
        ;

        private HWPFDocument u;
        private HWPFDocument a;
        [SetUp]
        public void SetUp()
        {
            u = HWPFTestDataSamples.OpenSampleFile("HeaderFooterUnicode.doc");
            a = HWPFTestDataSamples.OpenSampleFile("SampleDoc.doc");
        }

        [Test]
        public void TestAsciiTextParagraphs()
        {
            Range r = a.GetRange();
            Assert.AreEqual(
                    a_page_1 +
                    page_break + "\r" +
                    a_page_2,
                    r.Text
            );

            Assert.AreEqual(1, r.NumSections);
            Assert.AreEqual(1, a.SectionTable.GetSections().Count);
            Section s = r.GetSection(0);
            Assert.AreEqual(
                    a_page_1 +
                    page_break + "\r" +
                    a_page_2,
                    s.Text
            );

            Assert.AreEqual(
                    7,
                    r.NumParagraphs
            );
            String[] p1_parts = a_page_1.Split('\r');
            String[] p2_parts = a_page_2.Split('\r');

            // Check paragraph contents
            Assert.AreEqual(
                    p1_parts[0] + "\r",
                    r.GetParagraph(0).Text
            );
            Assert.AreEqual(
                    p1_parts[1] + "\r",
                    r.GetParagraph(1).Text
            );
            Assert.AreEqual(
                    p1_parts[2] + "\r",
                    r.GetParagraph(2).Text
            );

            Assert.AreEqual(
                    page_break + "\r",
                    r.GetParagraph(3).Text
            );

            Assert.AreEqual(
                    p2_parts[0] + "\r",
                    r.GetParagraph(4).Text
            );
            Assert.AreEqual(
                    p2_parts[1] + "\r",
                    r.GetParagraph(5).Text
            );
            Assert.AreEqual(
                    p2_parts[2] + "\r",
                    r.GetParagraph(6).Text
            );
        }
        [Test]
        public void TestAsciiStyling()
        {
            Range r = a.GetRange();

            Paragraph p1 = r.GetParagraph(0);
            Paragraph p7 = r.GetParagraph(6);

            Assert.AreEqual(1, p1.NumCharacterRuns);
            Assert.AreEqual(1, p7.NumCharacterRuns);

            CharacterRun c1 = p1.GetCharacterRun(0);
            CharacterRun c7 = p7.GetCharacterRun(0);

            Assert.AreEqual("Times New Roman", c1.GetFontName()); // No Calibri
            Assert.AreEqual("Arial Black", c7.GetFontName());
            Assert.AreEqual(20, c1.GetFontSize());
            Assert.AreEqual(32, c7.GetFontSize());

            // This document has 15 styles
            Assert.AreEqual(15, a.GetStyleSheet().NumStyles());

            // Ensure none of the paragraphs refer to one that isn't there,
            //  and none of their character Runs either
            for (int i = 0; i < a.GetRange().NumParagraphs; i++)
            {
                Paragraph p = a.GetRange().GetParagraph(i);
                Assert.IsTrue(p.GetStyleIndex() < 15);
            }
        }

        /**
         * Tests the raw defInitions of the paragraphs of
         *  a unicode document
         */
        [Test]
        public void TestUnicodeParagraphDefInitions()
        {
            Range r = u.GetRange();
            String[] p1_parts = u_page_1.Split('\r');
            String[] p2_parts = u_page_2.Split('\r');

            Assert.AreEqual(
                    u_page_1 + page_break + "\r" + u_page_2,
                    r.Text
            );
            Assert.AreEqual(
                    408, r.Text.Length
            );


            Assert.AreEqual(1, r.NumSections);
            Assert.AreEqual(1, u.SectionTable.GetSections().Count);
            Section s = r.GetSection(0);
            Assert.AreEqual(
                    u_page_1 +
                    page_break + "\r" +
                    u_page_2,
                    s.Text
            );
            Assert.AreEqual(0, s.StartOffset);
            Assert.AreEqual(408, s.EndOffset);


            List<PAPX> pDefs = r._paragraphs;
            Assert.AreEqual(35, pDefs.Count);

            // Check that the last paragraph ends where it should do
            Assert.AreEqual(531, u.GetOverallRange().Text.Length);
            Assert.AreEqual(530, u.GetCPSplitCalculator().GetHeaderTextboxEnd());
            PropertyNode pLast = (PropertyNode)pDefs[34];
            //		Assert.AreEqual(530, pLast.End);

            // Only care about the first few really though
            PropertyNode p0 = (PropertyNode)pDefs[0];
            PropertyNode p1 = (PropertyNode)pDefs[1];
            PropertyNode p2 = (PropertyNode)pDefs[2];
            PropertyNode p3 = (PropertyNode)pDefs[3];
            PropertyNode p4 = (PropertyNode)pDefs[4];

            // 5 paragraphs should get us to the end of our text
            Assert.IsTrue(p0.Start < 408);
            Assert.IsTrue(p0.End < 408);
            Assert.IsTrue(p1.Start < 408);
            Assert.IsTrue(p1.End < 408);
            Assert.IsTrue(p2.Start < 408);
            Assert.IsTrue(p2.End < 408);
            Assert.IsTrue(p3.Start < 408);
            Assert.IsTrue(p3.End < 408);
            Assert.IsTrue(p4.Start < 408);
            Assert.IsTrue(p4.End < 408);

            // Paragraphs should match with lines
            Assert.AreEqual(
                    0,
                    p0.Start
            );
            Assert.AreEqual(
                    p1_parts[0].Length + 1,
                    p0.End
            );

            Assert.AreEqual(
                    p1_parts[0].Length + 1,
                    p1.Start
            );
            Assert.AreEqual(
                    p1_parts[0].Length + 1 +
                    p1_parts[1].Length + 1,
                    p1.End
            );

            Assert.AreEqual(
                    p1_parts[0].Length + 1 +
                    p1_parts[1].Length + 1,
                    p2.Start
            );
            Assert.AreEqual(
                    p1_parts[0].Length + 1 +
                    p1_parts[1].Length + 1 +
                    p1_parts[2].Length + 1,
                    p2.End
            );
        }

        /**
         * Tests the paragraph text of a unicode document
         */
        [Test]
        public void TestUnicodeTextParagraphs()
        {
            Range r = u.GetRange();
            Assert.AreEqual(
                    u_page_1 +
                    page_break + "\r" +
                    u_page_2,
                    r.Text
            );

            Assert.AreEqual(
                    12,
                    r.NumParagraphs
            );
            String[] p1_parts = u_page_1.Split('\r');
            String[] p2_parts = u_page_2.Split('\r');

            // Check text all matches up properly
            Assert.AreEqual(p1_parts[0] + "\r", r.GetParagraph(0).Text);
            Assert.AreEqual(p1_parts[1] + "\r", r.GetParagraph(1).Text);
            Assert.AreEqual(p1_parts[2] + "\r", r.GetParagraph(2).Text);
            Assert.AreEqual(p1_parts[3] + "\r", r.GetParagraph(3).Text);
            Assert.AreEqual(p1_parts[4] + "\r", r.GetParagraph(4).Text);
            Assert.AreEqual(p1_parts[5] + "\r", r.GetParagraph(5).Text);
            Assert.AreEqual(p1_parts[6] + "\r", r.GetParagraph(6).Text);
            Assert.AreEqual(p1_parts[7] + "\r", r.GetParagraph(7).Text);
            Assert.AreEqual(p1_parts[8] + "\r", r.GetParagraph(8).Text);
            Assert.AreEqual(p1_parts[9] + "\r", r.GetParagraph(9).Text);
            Assert.AreEqual(page_break + "\r", r.GetParagraph(10).Text);
            Assert.AreEqual(p2_parts[0] + "\r", r.GetParagraph(11).Text);
        }
        [Test]
        public void TestUnicodeStyling()
        {
            Range r = u.GetRange();
            String[] p1_parts = u_page_1.Split('\r');

            Paragraph p1 = r.GetParagraph(0);
            Paragraph p7 = r.GetParagraph(6);

            // Line ending in its own run each time!
            Assert.AreEqual(2, p1.NumCharacterRuns);
            Assert.AreEqual(2, p7.NumCharacterRuns);

            CharacterRun c1a = p1.GetCharacterRun(0);
            CharacterRun c1b = p1.GetCharacterRun(1);
            CharacterRun c7a = p7.GetCharacterRun(0);
            CharacterRun c7b = p7.GetCharacterRun(1);

            Assert.AreEqual("Times New Roman", c1a.GetFontName()); // No Calibri
            Assert.AreEqual(22, c1a.GetFontSize());

            Assert.AreEqual("Times New Roman", c1b.GetFontName()); // No Calibri
            Assert.AreEqual(22, c1b.GetFontSize());

            Assert.AreEqual("Times New Roman", c7a.GetFontName());
            Assert.AreEqual(48, c7a.GetFontSize());

            Assert.AreEqual("Times New Roman", c7b.GetFontName());
            Assert.AreEqual(48, c7b.GetFontSize());

            // Now check where they crop up
            Assert.AreEqual(
                    0,
                    c1a.StartOffset
            );
            Assert.AreEqual(
                    p1_parts[0].Length,
                    c1a.EndOffset
            );

            Assert.AreEqual(
                    p1_parts[0].Length,
                    c1b.StartOffset
            );
            Assert.AreEqual(
                    p1_parts[0].Length + 1,
                    c1b.EndOffset
            );

            Assert.AreEqual(
                    p1_parts[0].Length + 1 +
                    p1_parts[1].Length + 1 +
                    p1_parts[2].Length + 1 +
                    p1_parts[3].Length + 1 +
                    p1_parts[4].Length + 1 +
                    p1_parts[5].Length + 1,
                    c7a.StartOffset
            );
            Assert.AreEqual(
                    p1_parts[0].Length + 1 +
                    p1_parts[1].Length + 1 +
                    p1_parts[2].Length + 1 +
                    p1_parts[3].Length + 1 +
                    p1_parts[4].Length + 1 +
                    p1_parts[5].Length + 1 +
                    1,
                    c7a.EndOffset
            );

            Assert.AreEqual(
                    p1_parts[0].Length + 1 +
                    p1_parts[1].Length + 1 +
                    p1_parts[2].Length + 1 +
                    p1_parts[3].Length + 1 +
                    p1_parts[4].Length + 1 +
                    p1_parts[5].Length + 1 +
                    1,
                    c7b.StartOffset
            );
            Assert.AreEqual(
                    p1_parts[0].Length + 1 +
                    p1_parts[1].Length + 1 +
                    p1_parts[2].Length + 1 +
                    p1_parts[3].Length + 1 +
                    p1_parts[4].Length + 1 +
                    p1_parts[5].Length + 1 +
                    p1_parts[6].Length + 1,
                    c7b.EndOffset
            );

            // This document has 15 styles
            Assert.AreEqual(15, a.GetStyleSheet().NumStyles());

            // Ensure none of the paragraphs refer to one that isn't there,
            //  and none of their character Runs either
            for (int i = 0; i < a.GetRange().NumParagraphs; i++)
            {
                Paragraph p = a.GetRange().GetParagraph(i);
                Assert.IsTrue(p.GetStyleIndex() < 15);
            }
        }
    }
}

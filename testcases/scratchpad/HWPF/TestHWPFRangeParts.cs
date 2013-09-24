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

namespace TestCases.HWPF
{
    
    using NPOI.HWPF;
    using System;
    using NPOI.HWPF.UserModel;
    using NUnit.Framework;



    /**
     * Test that we pull out the right bits of a file into
     *  the different ranges
     */
    [TestFixture]
    public class TestHWPFRangeParts
    {
        private static char page_break = (char)12;
        private static String headerDef =
            "\u0003\r\r" +
            "\u0004\r\r" +
            "\u0003\r\r" +
            "\u0004\r\r"
        ;
        private static String footerDef = "\r";
        private static String endHeaderFooter = "\r\r";


        private static String a_page_1 =
            "This is a sample word document. It has two pages. It has a three column heading, and a three column footer\r" +
            "\r" +
            "HEADING TEXT\r" +
            "\r" +
            "More on page one\r" +
            "\r\r" +
            "End of page 1\r"
        ;
        private static String a_page_2 =
            "This is page two. It also has a three column heading, and a three column footer.\r"
        ;

        private static String a_header =
            "First header column!\tMid header Right header!\r"
        ;
        private static String a_footer =
            "Footer Left\tFooter Middle Footer Right\r"
        ;


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

        private static String u_header =
            "\r\r" +
            "This is a simple header, with a \u20ac euro symbol in it.\r"
        ;
        private static String u_footer =
            "\r\r\r" +
            "The footer, with Moli\u00e8re, has Unicode in it.\r" +
            "\r\r\r\r"
        ;

        /**
         * A document made up only of basic ASCII text
         */
        private HWPFDocument docAscii;
        /**
         * A document with some unicode in it too
         */
        private HWPFDocument docUnicode;
        [SetUp]
        public void SetUp()
        {
            docUnicode = HWPFTestDataSamples.OpenSampleFile("HeaderFooterUnicode.doc");
            docAscii = HWPFTestDataSamples.OpenSampleFile("ThreeColHeadFoot.doc");
        }

        /**
         * Note - this Test Runs several times, to ensure that things
         *  don't get broken as we write out and read back in again
         * TODO - Make this work with 3+ Runs
         */
        [Test]
        public void TestBasics()
        {
            HWPFDocument doc = docAscii;
            for (int run = 0; run < 3; run++)
            {
                // First check the start and end bits
                Assert.AreEqual(
                        0,
                        doc._cpSplit.GetMainDocumentStart()
                );
                Assert.AreEqual(
                        a_page_1.Length +
                        2 + // page break
                        a_page_2.Length,
                        doc._cpSplit.GetMainDocumentEnd()
                );

                Assert.AreEqual(
                        238,
                        doc._cpSplit.GetFootnoteStart()
                );
                Assert.AreEqual(
                        238,
                        doc._cpSplit.GetFootnoteEnd()
                );

                Assert.AreEqual(
                        238,
                        doc._cpSplit.GetHeaderStoryStart()
                );
                Assert.AreEqual(
                        238 + headerDef.Length + a_header.Length +
                        footerDef.Length + a_footer.Length + endHeaderFooter.Length,
                        doc._cpSplit.GetHeaderStoryEnd()
                );

                // Write out and read back in again, Ready for
                //  the next run of the Test
                // TODO run more than once
                if (run < 1)
                    doc = HWPFTestDataSamples.WriteOutAndReadBack(doc);
            }
        }

        /**
         * Note - this Test Runs several times, to ensure that things
         *  don't get broken as we write out and read back in again
         * TODO - Make this work with 3+ Runs
         */
        [Test]
        public void TestContents()
        {
            HWPFDocument doc = docAscii;
            for (int run = 0; run < 3; run++)
            {
                Range r;

                // Now check the real ranges
                r = doc.GetRange();
                Assert.AreEqual(
                        a_page_1 +
                        page_break + "\r" +
                        a_page_2,
                        r.Text
                );

                r = doc.GetHeaderStoryRange();
                Assert.AreEqual(
                        headerDef +
                        a_header +
                        footerDef +
                        a_footer +
                        endHeaderFooter,
                        r.Text
                );

                r = doc.GetOverallRange();
                Assert.AreEqual(
                        a_page_1 +
                        page_break + "\r" +
                        a_page_2 +
                        headerDef +
                        a_header +
                        footerDef +
                        a_footer +
                        endHeaderFooter +
                        "\r",
                        r.Text
                );

                // Write out and read back in again, Ready for
                //  the next run of the Test
                // TODO run more than once
                if (run < 1)
                    doc = HWPFTestDataSamples.WriteOutAndReadBack(doc);
            }
        }

        /**
         * Note - this Test Runs several times, to ensure that things
         *  don't get broken as we write out and read back in again
         */
        [Test]
        public void TestBasicsUnicode()
        {
            HWPFDocument doc = docUnicode;
            for (int run = 0; run < 3; run++)
            {
                // First check the start and end bits
                Assert.AreEqual(
                        0,
                        doc._cpSplit.GetMainDocumentStart()
                );
                Assert.AreEqual(
                        u_page_1.Length +
                        2 + // page break
                        u_page_2.Length,
                        doc._cpSplit.GetMainDocumentEnd()
                );

                Assert.AreEqual(
                        408,
                        doc._cpSplit.GetFootnoteStart()
                );
                Assert.AreEqual(
                        408,
                        doc._cpSplit.GetFootnoteEnd()
                );

                Assert.AreEqual(
                        408,
                        doc._cpSplit.GetHeaderStoryStart()
                );
                // TODO - fix this one
                Assert.AreEqual(
                        408 + headerDef.Length + u_header.Length +
                        footerDef.Length + u_footer.Length + endHeaderFooter.Length,
                        doc._cpSplit.GetHeaderStoryEnd()
                );

                // Write out and read back in again, Ready for
                //  the next run of the Test
                // TODO run more than once
                if (run < 1)
                    doc = HWPFTestDataSamples.WriteOutAndReadBack(doc);
            }
        }
        [Test]
        public void TestContentsUnicode()
        {
            Range r;

            // Now check the real ranges
            r = docUnicode.GetRange();
            Assert.AreEqual(
                    u_page_1 +
                    page_break + "\r" +
                    u_page_2,
                    r.Text
            );

            r = docUnicode.GetHeaderStoryRange();
            Assert.AreEqual(
                    headerDef +
                    u_header +
                    footerDef +
                    u_footer +
                    endHeaderFooter,
                    r.Text
            );

            r = docUnicode.GetOverallRange();
            Assert.AreEqual(
                    u_page_1 +
                    page_break + "\r" +
                    u_page_2 +
                    headerDef +
                    u_header +
                    footerDef +
                    u_footer +
                    endHeaderFooter +
                    "\r",
                    r.Text
            );
        }
    }


}
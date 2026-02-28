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


using NUnit.Framework;
using TestCases.HWPF;
namespace NPOI.HWPF.UserModel
{

    /**
     * Tests for Word 6 and Word 95 support
     */
    [TestFixture]
    public class TestHWPFOldDocument : HWPFTestCase
    {
        /**
         * Test a simple Word 6 document
         */
        [Test]
        public void TestWord6()
        {
            // Can't open as HWPFDocument
            try
            {
                HWPFTestDataSamples.OpenSampleFile("Word6.doc");
                Assert.Fail("Shouldn't be openable");
            }
            catch (OldFileFormatException) { }

            // Open
            HWPFOldDocument doc = HWPFTestDataSamples.OpenOldSampleFile("Word6.doc");

            // Check
            Assert.AreEqual(1, doc.GetRange().NumSections);
            Assert.AreEqual(1, doc.GetRange().NumParagraphs);
            Assert.AreEqual(1, doc.GetRange().NumCharacterRuns);

            Assert.AreEqual(
                  "The quick brown fox jumps over the lazy dog\r",
                  doc.GetRange().GetParagraph(0).Text
            );
        }

        /**
         * Test a simple Word 95 document
         */
        [Test]
        public void TestWord95()
        {
            // Can't open as HWPFDocument
            try
            {
                HWPFTestDataSamples.OpenSampleFile("Word95.doc");
                Assert.Fail("Shouldn't be openable");
            }
            catch (OldFileFormatException) { }

            // Open
            HWPFOldDocument doc = HWPFTestDataSamples.OpenOldSampleFile("Word95.doc");

            // Check
            Assert.AreEqual(1, doc.GetRange().NumSections);
            Assert.AreEqual(7, doc.GetRange().NumParagraphs);

            Assert.AreEqual(
                  "The quick brown fox jumps over the lazy dog\r",
                  doc.GetRange().GetParagraph(0).Text
            );
            Assert.AreEqual("\r", doc.GetRange().GetParagraph(1).Text);
            Assert.AreEqual(
                  "Paragraph 2\r",
                  doc.GetRange().GetParagraph(2).Text
            );
            Assert.AreEqual("\r", doc.GetRange().GetParagraph(3).Text);
            Assert.AreEqual(
                  "Paragraph 3. Has some RED text and some " +
                  "BLUE BOLD text in it.\r",
                  doc.GetRange().GetParagraph(4).Text
            );
            Assert.AreEqual("\r", doc.GetRange().GetParagraph(5).Text);
            Assert.AreEqual(
                  "Last (4th) paragraph.\r",
                  doc.GetRange().GetParagraph(6).Text
            );

            Assert.AreEqual(1, doc.GetRange().GetParagraph(0).NumCharacterRuns);
            Assert.AreEqual(1, doc.GetRange().GetParagraph(1).NumCharacterRuns);
            Assert.AreEqual(1, doc.GetRange().GetParagraph(2).NumCharacterRuns);
            Assert.AreEqual(1, doc.GetRange().GetParagraph(3).NumCharacterRuns);
            // Normal, red, normal, blue+bold, normal
            Assert.AreEqual(5, doc.GetRange().GetParagraph(4).NumCharacterRuns);
            Assert.AreEqual(1, doc.GetRange().GetParagraph(5).NumCharacterRuns);
            // Normal, superscript for 4th, normal
            Assert.AreEqual(3, doc.GetRange().GetParagraph(6).NumCharacterRuns);
        }

        /**
         * Test a word document that has sections,
         *  as well as the usual paragraph stuff.
         */
        [Test]
        public void TestWord6Sections()
        {
            HWPFOldDocument doc = HWPFTestDataSamples.OpenOldSampleFile("Word6_sections.doc");

            Assert.AreEqual(3, doc.GetRange().NumSections);
            Assert.AreEqual(6, doc.GetRange().NumParagraphs);

            Assert.AreEqual(
                  "This is a Test.\r",
                  doc.GetRange().GetParagraph(0).Text
            );
            Assert.AreEqual("\r", doc.GetRange().GetParagraph(1).Text);
            Assert.AreEqual("\u000c", doc.GetRange().GetParagraph(2).Text); // Section line?
            Assert.AreEqual("This is a new section.\r", doc.GetRange().GetParagraph(3).Text);
            Assert.AreEqual("\u000c", doc.GetRange().GetParagraph(4).Text); // Section line?
            Assert.AreEqual("\r", doc.GetRange().GetParagraph(5).Text);
        }

        /**
         * Another word document with sections, this time with a 
         *  few more section properties set on it
         */
        [Test]
        public void TestWord6Sections2()
        {
            HWPFOldDocument doc = HWPFTestDataSamples.OpenOldSampleFile("Word6_sections2.doc");

            Assert.AreEqual(1, doc.GetRange().NumSections);
            Assert.AreEqual(57, doc.GetRange().NumParagraphs);

            Assert.AreEqual(
                  "\r",
                  doc.GetRange().GetParagraph(0).Text
            );
            Assert.AreEqual(
                  "STATEMENT  OF  INSOLVENCY  PRACTICE  10  (SCOTLAND)\r",
                  doc.GetRange().GetParagraph(1).Text
            );
        }
    }
}
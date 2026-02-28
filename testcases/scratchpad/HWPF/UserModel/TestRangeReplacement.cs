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

namespace TestCases.HWPF.UserModel
{
    
    using NPOI.HWPF.UserModel;
    using System;
    using NPOI.HWPF;
    using NUnit.Framework;

    /**
     *	Test to see if Range.ReplaceText() works even if the Range Contains a
     *	CharacterRun that uses Unicode characters.
     *
     * TODO - re-enable me when unicode paragraph stuff is fixed!
     */
    [TestFixture]
    public class TestRangeReplacement
    {

        // u201c and u201d are "smart-quotes"
        private String originalText =
            "It is used to confirm that text replacement works even if Unicode characters (such as \u201c\u2014\u201d (U+2014), \u201c\u2e8e\u201d (U+2E8E), or \u201c\u2714\u201d (U+2714)) are present.  Everybody should be thankful to the ${organization} and all the POI contributors for their assistance in this matter.\r";
        private String searchText = "${organization}";
        private String ReplacementText = "Apache Software Foundation";
        private String expectedText2 =
            "It is used to confirm that text replacement works even if Unicode characters (such as \u201c\u2014\u201d (U+2014), \u201c\u2e8e\u201d (U+2E8E), or \u201c\u2714\u201d (U+2714)) are present.  Everybody should be thankful to the Apache Software Foundation and all the POI contributors for their assistance in this matter.\r";
        private String expectedText3 = "Thank you, Apache Software Foundation!\r";

        private String illustrativeDocFile = "testRangeReplacement.doc";

        /**
         * Test just opening the files
         */
        [Test]
        public void TestOpen()
        {

            HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);
        }

        /**
         * Test (more "Confirm" than test) that we have the general structure that we expect to have.
         */
        [Test]
        public void TestDocStructure()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);

            Range range = daDoc.GetRange();
            Assert.AreEqual(414, range.Text.Length);

            Assert.AreEqual(1, range.NumSections);
            Section section = range.GetSection(0);
            Assert.AreEqual(414, section.Text.Length);

            Assert.AreEqual(5, section.NumParagraphs);
            Paragraph para = section.GetParagraph(2);

            Assert.AreEqual(5, para.NumCharacterRuns);
            String text =
                para.GetCharacterRun(0).Text +
                para.GetCharacterRun(1).Text +
                para.GetCharacterRun(2).Text +
                para.GetCharacterRun(3).Text +
                para.GetCharacterRun(4).Text
            ;

            Assert.AreEqual(originalText, text);
        }

        /**
         * Test that we can replace text in our Range with Unicode text.
         */
        [Test]
        public void TestRangeReplacementOne()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);

            // Has one section
            Range range = daDoc.GetRange();
            Assert.AreEqual(1, range.NumSections);

            // The first section has 5 paragraphs
            Section section = range.GetSection(0);
            Assert.AreEqual(5, section.NumParagraphs);


            // Change some text
            Paragraph para = section.GetParagraph(2);

            String text = para.Text;
            Assert.AreEqual(originalText, text);

            int offset = text.IndexOf(searchText);
            Assert.AreEqual(181, offset);

            para.ReplaceText(searchText, ReplacementText, offset);

            // Ensure we still have one section, 5 paragraphs
            Assert.AreEqual(1, range.NumSections);
            section = range.GetSection(0);

            Assert.AreEqual(5, section.NumParagraphs);
            para = section.GetParagraph(2);

            // Ensure the text is what we should now have
            text = para.Text;
            Assert.AreEqual(expectedText2, text);
        }

        /**
         * Test that we can replace text in our Range with Unicode text.
         */
        public void TestRangeReplacementAll()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);

            Range range = daDoc.GetRange();
            Assert.AreEqual(1, range.NumSections);

            Section section = range.GetSection(0);
            Assert.AreEqual(5, section.NumParagraphs);

            Paragraph para = section.GetParagraph(2);

            String text = para.Text;
            Assert.AreEqual(originalText, text);

            range.ReplaceText(searchText, ReplacementText);

            Assert.AreEqual(1, range.NumSections);
            section = range.GetSection(0);
            Assert.AreEqual(5, section.NumParagraphs);

            para = section.GetParagraph(2);
            text = para.Text;
            Assert.AreEqual(expectedText2, text);

            para = section.GetParagraph(3);
            text = para.Text;
            Assert.AreEqual(expectedText3, text);
        }
    }

}
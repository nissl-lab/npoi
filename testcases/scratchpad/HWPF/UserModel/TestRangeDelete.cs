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

using System;

using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.HWPF.Model;
using NUnit.Framework;
namespace TestCases.HWPF.UserModel
{

    /**
     *	Test to see if Range.Delete() works even if the Range Contains a
     *	CharacterRun that uses Unicode characters.
     */
    [TestFixture]
    public class TestRangeDelete
    {

        // u201c and u201d are "smart-quotes"
        private String introText =
            "Introduction\r";
        private String fillerText =
            "${delete} This is an MS-Word 97 formatted document created using NeoOffice v. 2.2.4 Patch 0 (OpenOffice.org v. 2.2.1).\r";
        private String originalText =
            "It is used to confirm that text delete works even if Unicode characters (such as \u201c\u2014\u201d (U+2014), \u201c\u2e8e\u201d (U+2E8E), or \u201c\u2714\u201d (U+2714)) are present.  Everybody should be thankful to the ${organization} ${delete} and all the POI contributors for their assistance in this matter.\r";
        private String lastText =
            "Thank you, ${organization} ${delete}!\r";
        private String searchText = "${delete}";
        private String expectedText1 = " This is an MS-Word 97 formatted document created using NeoOffice v. 2.2.4 Patch 0 (OpenOffice.org v. 2.2.1).\r";
        private String expectedText2 =
            "It is used to confirm that text delete works even if Unicode characters (such as \u201c\u2014\u201d (U+2014), \u201c\u2e8e\u201d (U+2E8E), or \u201c\u2714\u201d (U+2714)) are present.  Everybody should be thankful to the ${organization}  and all the POI contributors for their assistance in this matter.\r";
        private String expectedText3 = "Thank you, ${organization} !\r";

        private String illustrativeDocFile;

        [SetUp]
        public void SetUp()
        {
            illustrativeDocFile = "testRangeDelete.doc";
        }

        /**
         * Test just opening the files
         */
        [Test]
        public void TestOpen()
        {

            HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);
        }

        /**
         * Test (more "Confirm" than Test) that we have the general structure that we expect to have.
         */
        [Test]
        public void TestDocStructure()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);
            Range range;
            Section section;
            Paragraph para;
            PAPX paraDef;

            // First, check overall
            range = daDoc.GetOverallRange();
            Assert.AreEqual(1, range.NumSections);
            Assert.AreEqual(5, range.NumParagraphs);


            // Now, onto just the doc bit
            range = daDoc.GetRange();

            Assert.AreEqual(1, range.NumSections);
            Assert.AreEqual(1, daDoc.SectionTable.GetSections().Count);
            section = range.GetSection(0);

            Assert.AreEqual(5, section.NumParagraphs);

            para = section.GetParagraph(0);
            Assert.AreEqual(1, para.NumCharacterRuns);
            Assert.AreEqual(introText, para.Text);

            para = section.GetParagraph(1);
            Assert.AreEqual(5, para.NumCharacterRuns);
            Assert.AreEqual(fillerText, para.Text);


            paraDef = (PAPX)daDoc.ParagraphTable.GetParagraphs()[2];
            Assert.AreEqual(132, paraDef.Start);
            Assert.AreEqual(400, paraDef.End);

            para = section.GetParagraph(2);
            Assert.AreEqual(5, para.NumCharacterRuns);
            Assert.AreEqual(originalText, para.Text);


            paraDef = (PAPX)daDoc.ParagraphTable.GetParagraphs()[3];
            Assert.AreEqual(400, paraDef.Start);
            Assert.AreEqual(438, paraDef.End);

            para = section.GetParagraph(3);
            Assert.AreEqual(1, para.NumCharacterRuns);
            Assert.AreEqual(lastText, para.Text);


            // Check things match on text length
            Assert.AreEqual(439, range.Text.Length);
            Assert.AreEqual(439, section.Text.Length);
            Assert.AreEqual(439,
                    section.GetParagraph(0).Text.Length +
                    section.GetParagraph(1).Text.Length +
                    section.GetParagraph(2).Text.Length +
                    section.GetParagraph(3).Text.Length +
                    section.GetParagraph(4).Text.Length
            );
        }

        /**
         * Test that we can delete text (one instance) from our Range with Unicode text.
         */
        [Test]
        public void TestRangeDeleteOne()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);

            Range range = daDoc.GetOverallRange();
            Assert.AreEqual(1, range.NumSections);

            Section section = range.GetSection(0);
            Assert.AreEqual(5, section.NumParagraphs);

            Paragraph para = section.GetParagraph(2);

            String text = para.Text;
            Assert.AreEqual(originalText, text);

            int offset = text.IndexOf(searchText);
            Assert.AreEqual(192, offset);

            int absOffset = para.StartOffset + offset;
            Range subRange = new Range(absOffset, (absOffset + searchText.Length), para.GetDocument());

            Assert.AreEqual(searchText, subRange.Text);

            subRange.Delete();

            // we need to let the model re-calculate the Range before we Evaluate it
            range = daDoc.GetRange();

            Assert.AreEqual(1, range.NumSections);
            section = range.GetSection(0);

            Assert.AreEqual(5, section.NumParagraphs);
            para = section.GetParagraph(2);

            text = para.Text;
            Assert.AreEqual(expectedText2, text);

            // this can lead to a StringBuilderOutOfBoundsException, so we will add it
            // even though we don't have an assertion for it
            Range daRange = daDoc.GetRange();
            text = daRange.Text;
        }

        /**
         * Test that we can delete text (all instances of) from our Range with Unicode text.
         */
        [Test]
        public void TestRangeDeleteAll()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);

            Range range = daDoc.GetRange();
            Assert.AreEqual(1, range.NumSections);

            Section section = range.GetSection(0);
            Assert.AreEqual(5, section.NumParagraphs);

            Paragraph para = section.GetParagraph(2);

            String text = para.Text;
            Assert.AreEqual(originalText, text);

            bool keepLooking = true;
            while (keepLooking)
            {
                // Reload the range every time
                range = daDoc.GetRange();
                int offset = range.Text.IndexOf(searchText);
                if (offset >= 0)
                {

                    int absOffset = range.StartOffset + offset;

                    Range subRange = new Range(
                        absOffset, (absOffset + searchText.Length), range.GetDocument());

                    Assert.AreEqual(searchText, subRange.Text);

                    subRange.Delete();

                }
                else
                {
                    keepLooking = false;
                }
            }

            // we need to let the model re-calculate the Range before we use it
            range = daDoc.GetRange();

            Assert.AreEqual(1, range.NumSections);
            section = range.GetSection(0);

            Assert.AreEqual(5, section.NumParagraphs);

            para = section.GetParagraph(0);
            text = para.Text;
            Assert.AreEqual(introText, text);

            para = section.GetParagraph(1);
            text = para.Text;
            Assert.AreEqual(expectedText1, text);

            para = section.GetParagraph(2);
            text = para.Text;
            Assert.AreEqual(expectedText2, text);

            para = section.GetParagraph(3);
            text = para.Text;
            Assert.AreEqual(expectedText3, text);
        }
    }
}


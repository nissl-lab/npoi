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

using TestCases.HWPF;
using System;
using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NUnit.Framework;

namespace TestCases.HWPF.UserModel
{

    /**
     *	Test to see if Range.insertBefore() works even if the Range Contains a
     *	CharacterRun that uses Unicode characters.
     *
     * TODO - re-enable me when unicode paragraph stuff is fixed!
     */
    [TestFixture]
    public class TestRangeInsertion
    {

        // u201c and u201d are "smart-quotes"
        private String originalText =
            "It is used to confirm that text insertion works even if Unicode characters (such as \u201c\u2014\u201d (U+2014), \u201c\u2e8e\u201d (U+2E8E), or \u201c\u2714\u201d (U+2714)) are present.\r";
        private String textToInsert = "Look at me!  I'm cool!  ";
        private int insertionPoint = 122;

        private String illustrativeDocFile;

        [SetUp]
        public void SetUp()
        {
            illustrativeDocFile = "testRangeInsertion.doc";
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
         * Test (more "Confirm" than test) that we have the general structure that we expect to have.
         */
        [Test]
        public void TestDocStructure()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);

            Range range = daDoc.GetRange();

            Assert.AreEqual(1, range.NumSections);
            Section section = range.GetSection(0);

            Assert.AreEqual(3, section.NumParagraphs);
            Paragraph para = section.GetParagraph(2);
            Assert.AreEqual(originalText, para.Text);

            Assert.AreEqual(3, para.NumCharacterRuns);
            String text =
                para.GetCharacterRun(0).Text +
                para.GetCharacterRun(1).Text +
                para.GetCharacterRun(2).Text
            ;

            Assert.AreEqual(originalText, text);

            Assert.AreEqual(insertionPoint, para.StartOffset);
        }

        /**
         * Test that we can insert text in our CharacterRun with Unicode text.
         */
        [Test]
        public void TestRangeInsert()
        {

            HWPFDocument daDoc = HWPFTestDataSamples.OpenSampleFile(illustrativeDocFile);

            //if (false)
            //{ // TODO - delete or resurrect this code
            //    Range range1 = daDoc.GetRange();
            //    Section section1 = range1.GetSection(0);
            //    Paragraph para1 = section1.GetParagraph(2);
            //    String text1 = para1.GetCharacterRun(0).Text + para1.GetCharacterRun(1).Text +
            //    para1.GetCharacterRun(2).Text;

            //    Console.WriteLine(text1);
            //}

            Range range = new Range(insertionPoint, (insertionPoint + 2), daDoc);
            range.InsertBefore(textToInsert);

            // we need to let the model re-calculate the Range before we Evaluate it
            range = daDoc.GetRange();

            Assert.AreEqual(1, range.NumSections);
            Section section = range.GetSection(0);

            Assert.AreEqual(3, section.NumParagraphs);
            Paragraph para = section.GetParagraph(2);
            Assert.AreEqual((textToInsert + originalText), para.Text);

            Assert.AreEqual(3, para.NumCharacterRuns);
            String text =
                para.GetCharacterRun(0).Text +
                para.GetCharacterRun(1).Text +
                para.GetCharacterRun(2).Text
            ;

            // Console.WriteLine(text);

            Assert.AreEqual((textToInsert + originalText), text);
        }
    }
}

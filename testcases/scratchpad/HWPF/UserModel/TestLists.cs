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

using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NUnit.Framework;

namespace TestCases.HWPF.UserModel
{
    /**
     * Tests for our handling of lists
     */
    [TestFixture]
    public class TestLists
    {
        [Test]
        public void TestBasics()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Lists.doc");
            Range r = doc.GetRange();

            Assert.AreEqual(40, r.NumParagraphs);
            Assert.AreEqual("Heading Level 1\r", r.GetParagraph(0).Text);
            Assert.AreEqual("This document has different lists in it for testing\r", r.GetParagraph(1).Text);
            Assert.AreEqual("The end!\r", r.GetParagraph(38).Text);
            Assert.AreEqual("\r", r.GetParagraph(39).Text);

            Assert.AreEqual(9, r.GetParagraph(0).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(1).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(38).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(39).GetLvl());
        }
        [Test]
        public void TestUnorderedLists()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Lists.doc");
            Range r = doc.GetRange();
            Assert.AreEqual(40, r.NumParagraphs);

            // Normal bullet points
            Assert.AreEqual("This document has different lists in it for testing\r", r.GetParagraph(1).Text);
            Assert.AreEqual("Unordered list 1\r", r.GetParagraph(2).Text);
            Assert.AreEqual("UL 2\r", r.GetParagraph(3).Text);
            Assert.AreEqual("UL 3\r", r.GetParagraph(4).Text);
            Assert.AreEqual("Next up is an ordered list:\r", r.GetParagraph(5).Text);

            Assert.AreEqual(9, r.GetParagraph(1).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(2).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(3).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(4).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(5).GetLvl());

            Assert.AreEqual(0, r.GetParagraph(1).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(2).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(3).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(4).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(5).GetIlvl());

            // Tick bullets
            Assert.AreEqual("Now for an un-ordered list with a different bullet style:\r", r.GetParagraph(9).Text);
            Assert.AreEqual("Tick 1\r", r.GetParagraph(10).Text);
            Assert.AreEqual("Tick 2\r", r.GetParagraph(11).Text);
            Assert.AreEqual("Multi-level un-ordered list:\r", r.GetParagraph(12).Text);

            Assert.AreEqual(9, r.GetParagraph(9).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(10).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(11).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(12).GetLvl());

            Assert.AreEqual(0, r.GetParagraph(9).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(10).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(11).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(12).GetIlvl());

            // TODO Test for tick not bullet
        }
        [Test]
        public void TestOrderedLists()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Lists.doc");
            Range r = doc.GetRange();
            Assert.AreEqual(40, r.NumParagraphs);

            Assert.AreEqual("Next up is an ordered list:\r", r.GetParagraph(5).Text);
            Assert.AreEqual("Ordered list 1\r", r.GetParagraph(6).Text);
            Assert.AreEqual("OL 2\r", r.GetParagraph(7).Text);
            Assert.AreEqual("OL 3\r", r.GetParagraph(8).Text);
            Assert.AreEqual("Now for an un-ordered list with a different bullet style:\r", r.GetParagraph(9).Text);

            Assert.AreEqual(9, r.GetParagraph(5).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(6).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(7).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(8).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(9).GetLvl());

            Assert.AreEqual(0, r.GetParagraph(5).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(6).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(7).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(8).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(9).GetIlvl());
        }
        [Test]
        public void TestMultiLevelLists()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Lists.doc");
            Range r = doc.GetRange();
            Assert.AreEqual(40, r.NumParagraphs);

            Assert.AreEqual("Multi-level un-ordered list:\r", r.GetParagraph(12).Text);
            Assert.AreEqual("ML 1:1\r", r.GetParagraph(13).Text);
            Assert.AreEqual("ML 1:2\r", r.GetParagraph(14).Text);
            Assert.AreEqual("ML 2:1\r", r.GetParagraph(15).Text);
            Assert.AreEqual("ML 2:2\r", r.GetParagraph(16).Text);
            Assert.AreEqual("ML 2:3\r", r.GetParagraph(17).Text);
            Assert.AreEqual("ML 3:1\r", r.GetParagraph(18).Text);
            Assert.AreEqual("ML 4:1\r", r.GetParagraph(19).Text);
            Assert.AreEqual("ML 5:1\r", r.GetParagraph(20).Text);
            Assert.AreEqual("ML 5:2\r", r.GetParagraph(21).Text);
            Assert.AreEqual("ML 2:4\r", r.GetParagraph(22).Text);
            Assert.AreEqual("ML 1:3\r", r.GetParagraph(23).Text);
            Assert.AreEqual("Multi-level ordered list:\r", r.GetParagraph(24).Text);
            Assert.AreEqual("OL 1\r", r.GetParagraph(25).Text);
            Assert.AreEqual("OL 2\r", r.GetParagraph(26).Text);
            Assert.AreEqual("OL 2.1\r", r.GetParagraph(27).Text);
            Assert.AreEqual("OL 2.2\r", r.GetParagraph(28).Text);
            Assert.AreEqual("OL 2.2.1\r", r.GetParagraph(29).Text);
            Assert.AreEqual("OL 2.2.2\r", r.GetParagraph(30).Text);
            Assert.AreEqual("OL 2.2.2.1\r", r.GetParagraph(31).Text);
            Assert.AreEqual("OL 2.2.3\r", r.GetParagraph(32).Text);
            Assert.AreEqual("OL 3\r", r.GetParagraph(33).Text);
            Assert.AreEqual("Finally we want some indents, to tell the difference\r", r.GetParagraph(34).Text);

            for (int i = 12; i <= 34; i++)
            {
                Assert.AreEqual(9, r.GetParagraph(12).GetLvl());
            }
            Assert.AreEqual(0, r.GetParagraph(12).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(13).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(14).GetIlvl());
            Assert.AreEqual(1, r.GetParagraph(15).GetIlvl());
            Assert.AreEqual(1, r.GetParagraph(16).GetIlvl());
            Assert.AreEqual(1, r.GetParagraph(17).GetIlvl());
            Assert.AreEqual(2, r.GetParagraph(18).GetIlvl());
            Assert.AreEqual(3, r.GetParagraph(19).GetIlvl());
            Assert.AreEqual(4, r.GetParagraph(20).GetIlvl());
            Assert.AreEqual(4, r.GetParagraph(21).GetIlvl());
            Assert.AreEqual(1, r.GetParagraph(22).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(23).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(24).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(25).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(26).GetIlvl());
            Assert.AreEqual(1, r.GetParagraph(27).GetIlvl());
            Assert.AreEqual(1, r.GetParagraph(28).GetIlvl());
            Assert.AreEqual(2, r.GetParagraph(29).GetIlvl());
            Assert.AreEqual(2, r.GetParagraph(30).GetIlvl());
            Assert.AreEqual(3, r.GetParagraph(31).GetIlvl());
            Assert.AreEqual(2, r.GetParagraph(32).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(33).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(34).GetIlvl());
        }
        [Test]
        public void TestIndentedText()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Lists.doc");
            Range r = doc.GetRange();

            Assert.AreEqual(40, r.NumParagraphs);
            Assert.AreEqual("Finally we want some indents, to tell the difference\r", r.GetParagraph(34).Text);
            Assert.AreEqual("Indented once\r", r.GetParagraph(35).Text);
            Assert.AreEqual("Indented twice\r", r.GetParagraph(36).Text);
            Assert.AreEqual("Indented three times\r", r.GetParagraph(37).Text);
            Assert.AreEqual("The end!\r", r.GetParagraph(38).Text);

            Assert.AreEqual(9, r.GetParagraph(34).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(35).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(36).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(37).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(38).GetLvl());
            Assert.AreEqual(9, r.GetParagraph(39).GetLvl());

            Assert.AreEqual(0, r.GetParagraph(34).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(35).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(36).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(37).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(38).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(39).GetIlvl());

            // TODO Test the indent
        }
        [Test]
        public void TestWriteRead()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Lists.doc");
            doc = HWPFTestDataSamples.WriteOutAndReadBack(doc);

            Range r = doc.GetRange();

            // Check a couple at random
            Assert.AreEqual(4, r.GetParagraph(21).GetIlvl());
            Assert.AreEqual(1, r.GetParagraph(22).GetIlvl());
            Assert.AreEqual(0, r.GetParagraph(23).GetIlvl());
        }
    }

}
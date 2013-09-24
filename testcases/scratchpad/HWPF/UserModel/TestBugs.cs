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
    using NPOI.HWPF;
    using NPOI.HWPF.UserModel;
    using NPOI.HWPF.Model;
    
    using System;
    using NUnit.Framework;

    [TestFixture]
    public class TestBugs
    {
        [Test]
        public void Test50075()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug50075.doc");
            Range range = doc.GetRange();
            Assert.AreEqual(1, range.NumParagraphs);
            Paragraph para1 = (Paragraph)range.GetParagraph(0);
            ListEntry entry = new ListEntry(para1._paragraphs[0], (Range)para1._parent,doc.GetListTables());
            ListFormatOverride override1 = doc.GetListTables().GetOverride(entry.GetIlfo());
            ListLevel level = doc.GetListTables().GetLevel(override1.GetLsid(), entry.GetIlvl());

            // the bug reproduces, if this call fails with NullPointerException
            level.GetNumberText();
        }
        [Test]
        public void Test49820()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug49820.doc");

            Range documentRange = doc.GetRange();
            StyleSheet styleSheet = doc.GetStyleSheet();

            // JUnit asserts
            assertLevels(documentRange, styleSheet, 0, 0, 0);
            assertLevels(documentRange, styleSheet, 1, 1, 1);
            assertLevels(documentRange, styleSheet, 2, 2, 2);
            assertLevels(documentRange, styleSheet, 3, 3, 3);
            assertLevels(documentRange, styleSheet, 4, 4, 4);
            assertLevels(documentRange, styleSheet, 5, 5, 5);
            assertLevels(documentRange, styleSheet, 6, 6, 6);
            assertLevels(documentRange, styleSheet, 7, 7, 7);
            assertLevels(documentRange, styleSheet, 8, 8, 8);
            assertLevels(documentRange, styleSheet, 9, 9, 9);
            assertLevels(documentRange, styleSheet, 10, 9, 0);
            assertLevels(documentRange, styleSheet, 11, 9, 4);

            // output to console
            for (int i=0; i<documentRange.NumParagraphs; i++) {
              Paragraph par = documentRange.GetParagraph(i);
              int styleLvl = styleSheet.GetParagraphStyle(par.GetStyleIndex()).GetLvl();
              int parLvl = par.GetLvl();
              Console.WriteLine("Style level: " + styleLvl + ", paragraph level: " + parLvl + ", text: " + par.Text);
            }
        }

        private void assertLevels(Range documentRange, StyleSheet styleSheet, int parIndex, int expectedStyleLvl, int expectedParLvl)
        {
            Paragraph par = documentRange.GetParagraph(parIndex);
            Assert.AreEqual(expectedStyleLvl, styleSheet.GetParagraphStyle(par.GetStyleIndex()).GetLvl());
            Assert.AreEqual(expectedParLvl, par.GetLvl());
        }
    }
}


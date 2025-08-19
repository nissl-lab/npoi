/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XSSF.Extractor
{

    using NPOI;
    using NPOI.HSSF;
    using NPOI.HSSF.Extractor;
    using NPOI.XSSF;
    using NPOI.XSSF.Extractor;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System.Text.RegularExpressions;
    using TestCases.HSSF;

    /// <summary>
    /// Tests for <see cref="XSSFEventBasedExcelExtractor"/>
    /// </summary>
    [TestFixture]
    public class TestXSSFEventBasedExcelExtractor
    {
        protected XSSFEventBasedExcelExtractor GetExtractor(String sampleName)
        {

            return new XSSFEventBasedExcelExtractor(XSSFTestDataSamples.OpenSamplePackage(sampleName));
        }

        /// <summary>
        /// Get text out of the simple file
        /// </summary>
        [Test]
        public void TestGetSimpleText()
        {

            // a very simple file
            XSSFEventBasedExcelExtractor extractor = GetExtractor("sample.xlsx");
            var _ = extractor.Text;

            String text = extractor.Text;
            ClassicAssert.IsTrue(text.Length > 0);

            // Check sheet names
            POITestCase.AssertStartsWith(text, "Sheet1");
            POITestCase.AssertEndsWith(text, "Sheet3\n");

            // Now without, will have text
            extractor.IncludeSheetNames = (false);
            text = extractor.Text;
            String CHUNK1 =
            "Lorem\t111\n" +
            "ipsum\t222\n" +
            "dolor\t333\n" +
            "sit\t444\n" +
            "amet\t555\n" +
            "consectetuer\t666\n" +
            "adipiscing\t777\n" +
            "elit\t888\n" +
            "Nunc\t999\n";
            String CHUNK2 =
            "The quick brown fox jumps over the lazy dog\n" +
            "hello, xssf	hello, xssf\n" +
            "hello, xssf	hello, xssf\n" +
            "hello, xssf	hello, xssf\n" +
            "hello, xssf	hello, xssf\n";
            ClassicAssert.AreEqual(
                    CHUNK1 +
                    "at\t4995\n" +
                    CHUNK2
                    , text);

            // Now Get formulas not their values
            extractor.FormulasNotResults = (true);
            text = extractor.Text;
            ClassicAssert.AreEqual(
                    CHUNK1 +
                    "at\tSUM(B1:B9)\n" +
                    CHUNK2, text);

            // With sheet names too
            extractor.IncludeSheetNames = (true);
            text = extractor.Text;
            ClassicAssert.AreEqual(
                    "Sheet1\n" +
                    CHUNK1 +
                    "at\tSUM(B1:B9)\n" +
                    "rich test\n" +
                    CHUNK2 +
                    "Sheet3\n"
                    , text);

            extractor.Close();
        }

        [Test]
        public void TestGetComplexText()
        {

            // A fairly complex file
            XSSFEventBasedExcelExtractor extractor = GetExtractor("AverageTaxRates.xlsx");
            var _ = extractor.Text;

            String text = extractor.Text;
            ClassicAssert.IsTrue(text.Length > 0);

            // Might not have all formatting it should do!
            POITestCase.AssertStartsWith(text,
                            "Avgtxfull\n" +
                            "(iii) AVERAGE TAX RATES ON ANNUAL"
            );

            extractor.Close();
        }

        [Test]
        public void TestInlineStrings()
        {

            XSSFEventBasedExcelExtractor extractor = GetExtractor("InlineStrings.xlsx");
            extractor.FormulasNotResults = (true);
            String text = extractor.Text;

            // Numbers
            POITestCase.AssertContains(text, "43");
            POITestCase.AssertContains(text, "22");

            // Strings
            POITestCase.AssertContains(text, "ABCDE");
            POITestCase.AssertContains(text, "Long Text");

            // Inline Strings
            POITestCase.AssertContains(text, "1st Inline String");
            POITestCase.AssertContains(text, "And More");

            // Formulas
            POITestCase.AssertContains(text, "A2");
            POITestCase.AssertContains(text, "A5-A$2");

            extractor.Close();
        }

        /// <summary>
        /// Test that we return pretty much the same as
        ///  ExcelExtractor does, when we're both passed
        ///  the same file, just saved as xls and xlsx
        /// </summary>
        [Test]
        public void TestComparedToOLE2()
        {

            // A fairly simple file - ooxml
            XSSFEventBasedExcelExtractor ooxmlExtractor = GetExtractor("SampleSS.xlsx");

            ExcelExtractor ole2Extractor =
            new ExcelExtractor(HSSFTestDataSamples.OpenSampleWorkbook("SampleSS.xls"));

            POITextExtractor[] extractors =
            new POITextExtractor[] { ooxmlExtractor, ole2Extractor };
            foreach(POITextExtractor extractor in extractors)
            {
                String text = extractor.Text.Replace("\r", "").Replace("\t", "");
                POITestCase.AssertStartsWith(text, "First Sheet\nTest spreadsheet\n2nd row2nd row 2nd column\n");
                Regex pattern = new Regex(".*13(\\.0+)?\\s+Sheet3.*", RegexOptions.Compiled | RegexOptions.Singleline);
                Match m = pattern.Match(text);
                ClassicAssert.IsTrue(m.Success);
            }

            ole2Extractor.Close();
            ooxmlExtractor.Close();
        }

        /// <summary>
        /// Test text extraction from text box using GetShapes()
        /// </summary>
        /// <exception cref="Exception">Exception</exception>
        [Test]
        public void TestShapes()
        {
            XSSFEventBasedExcelExtractor ooxmlExtractor = GetExtractor("WithTextBox.xlsx");
            try
            {
                String text = ooxmlExtractor.Text;
                StringAssert.Contains("Line 1", text);
                StringAssert.Contains("Line 2", text);
                StringAssert.Contains("Line 3", text);
            }
            finally
            {
                ooxmlExtractor.Close();
            }
        }

        /// <summary>
        /// Test that we return the same output for unstyled numbers as the
        /// non-event-based XSSFExcelExtractor.
        /// </summary>
        [Test]
        public void TestUnstyledNumbersComparedToNonEventBasedExtractor()
        {
            String expectedOutput = "Sheet1\n99.99\n";
            XSSFExcelExtractor extractor = new XSSFExcelExtractor(
                XSSFTestDataSamples.OpenSampleWorkbook("56011.xlsx"));
            try
            {
                ClassicAssert.AreEqual(expectedOutput, extractor.Text.Replace(",", "."));
            }
            finally
            {
                extractor.Close();
            }

            XSSFEventBasedExcelExtractor fixture =
                new XSSFEventBasedExcelExtractor(
                        XSSFTestDataSamples.OpenSamplePackage("56011.xlsx"));
            try
            {
                ClassicAssert.AreEqual(expectedOutput, fixture.Text.Replace(",", "."));
            }
            finally
            {
                fixture.Close();
            }
        }

        /// <summary>
        /// Test that we return the same output headers and footers as the
        /// non-event-based XSSFExcelExtractor.
        /// </summary>
        [Test]
        public void TestHeadersAndFootersComparedToNonEventBasedExtractor()
        {
            String expectedOutputWithHeadersAndFooters =
                "Sheet1\n" +
                "&\"Calibri,Regular\"&K000000top left\t&\"Calibri,Regular\"&K000000top center\t&\"Calibri,Regular\"&K000000top right\n" +
                "abc\t123\n" +
                "&\"Calibri,Regular\"&K000000bottom left\t&\"Calibri,Regular\"&K000000bottom center\t&\"Calibri,Regular\"&K000000bottom right\n";

            String expectedOutputWithoutHeadersAndFooters =
                "Sheet1\n" +
                "abc\t123\n";

            XSSFExcelExtractor extractor = new XSSFExcelExtractor(
                XSSFTestDataSamples.OpenSampleWorkbook("headerFooterTest.xlsx"));
            try
            {
                ClassicAssert.AreEqual(expectedOutputWithHeadersAndFooters, extractor.Text);
                extractor.IncludeHeadersFooters = (false);
                ClassicAssert.AreEqual(expectedOutputWithoutHeadersAndFooters, extractor.Text);
            }
            finally
            {
                extractor.Close();
            }

            XSSFEventBasedExcelExtractor fixture =
                new XSSFEventBasedExcelExtractor(
                        XSSFTestDataSamples.OpenSamplePackage("headerFooterTest.xlsx"));
            try
            {
                ClassicAssert.AreEqual(expectedOutputWithHeadersAndFooters, fixture.Text);
                fixture.IncludeHeadersFooters = (false);
                ClassicAssert.AreEqual(expectedOutputWithoutHeadersAndFooters, fixture.Text);
            }
            finally
            {
                fixture.Close();
            }
        }

        /// <summary>
        /// <para>
        /// Test that XSSFEventBasedExcelExtractor outputs comments when specified.
        /// The output will contain two improvements over the output from
        ///  XSSFExcelExtractor in that (1) comments from empty cells will be
        /// outputted, and (2) the author will not be outputted twice.
        /// </para>
        /// <para>
        /// This test will need to be modified if these improvements are ported to
        /// XSSFExcelExtractor.
        /// </para>
        /// </summary>
        [Test]
        public void TestCommentsComparedToNonEventBasedExtractor()
        {
            String expectedOutputWithoutComments =
                "Sheet1\n" +
                "\n" +
                "abc\n" +
                "\n" +
                "123\n" +
                "\n" +
                "\n" +
                "\n";

            String nonEventBasedExtractorOutputWithComments =
                "Sheet1\n" +
                "\n" +
                "abc Comment by Shaun Kalley: Shaun Kalley: Comment A2\n" +
                "\n" +
                "123 Comment by Shaun Kalley: Shaun Kalley: Comment B4\n" +
                "\n" +
                "\n" +
                "\n";

            String eventBasedExtractorOutputWithComments =
                "Sheet1\n" +
                "Comment by Shaun Kalley: Comment A1\tComment by Shaun Kalley: Comment B1\n" +
                "abc Comment by Shaun Kalley: Comment A2\tComment by Shaun Kalley: Comment B2\n" +
                "Comment by Shaun Kalley: Comment A3\tComment by Shaun Kalley: Comment B3\n" +
                "Comment by Shaun Kalley: Comment A4\t123 Comment by Shaun Kalley: Comment B4\n" +
                "Comment by Shaun Kalley: Comment A5\tComment by Shaun Kalley: Comment B5\n" +
                "Comment by Shaun Kalley: Comment A7\tComment by Shaun Kalley: Comment B7\n" +
                "Comment by Shaun Kalley: Comment A8\tComment by Shaun Kalley: Comment B8\n";

            XSSFExcelExtractor extractor = new XSSFExcelExtractor(
                XSSFTestDataSamples.OpenSampleWorkbook("commentTest.xlsx"));
            try
            {
                extractor.AddTabEachEmptyCell = false;
                ClassicAssert.AreEqual(expectedOutputWithoutComments, extractor.Text);
                extractor.IncludeCellComments = (true);
                ClassicAssert.AreEqual(nonEventBasedExtractorOutputWithComments, extractor.Text);
            }
            finally
            {
                extractor.Close();
            }

            XSSFEventBasedExcelExtractor fixture =
                new XSSFEventBasedExcelExtractor(
                        XSSFTestDataSamples.OpenSamplePackage("commentTest.xlsx"));
            try
            {
                ClassicAssert.AreEqual(expectedOutputWithoutComments, fixture.Text);
                fixture.IncludeCellComments = (true);
                ClassicAssert.AreEqual(eventBasedExtractorOutputWithComments, fixture.Text);
            }
            finally
            {
                fixture.Close();
            }
        }

        [Test]
        public void TestFile56278_normal()
        {

            // first with normal Text Extractor
            POIXMLTextExtractor extractor = new XSSFExcelExtractor(
                XSSFTestDataSamples.OpenSampleWorkbook("56278.xlsx"));
            try
            {
                ClassicAssert.IsNotNull(extractor.Text);
            }
            finally
            {
                extractor.Close();
            }
        }

        [Test]
        public void TestFile56278_event()
        {

            // then with event based one
            POIXMLTextExtractor extractor = GetExtractor("56278.xlsx");
            try
            {
                ClassicAssert.IsNotNull(extractor.Text);
            }
            finally
            {
                extractor.Close();
            }
        }

        [Test]
        public void Test59021()
        {

            XSSFEventBasedExcelExtractor ex =
                new XSSFEventBasedExcelExtractor(
                        XSSFTestDataSamples.OpenSamplePackage("59021.xlsx"));
            String text = ex.Text;
            StringAssert.Contains("Abkhazia - Fixed", text);
            StringAssert.Contains("10/02/2016", text);
            ex.Close();
        }

        [Test]
        public void Test51519()
        {

            //default behavior: include phonetic runs
            XSSFEventBasedExcelExtractor ex =
                new XSSFEventBasedExcelExtractor(
                        XSSFTestDataSamples.OpenSamplePackage("51519.xlsx"));
            String text = ex.Text;
            StringAssert.Contains("\u65E5\u672C\u30AA\u30E9\u30AF\u30EB \u30CB\u30DB\u30F3", text);
            ex.Close();

            //now try turning them off
            ex = new XSSFEventBasedExcelExtractor(
                            XSSFTestDataSamples.OpenSamplePackage("51519.xlsx"));
            ex.SetConcatenatePhoneticRuns(false);
            text = ex.Text;
            ClassicAssert.IsFalse(text.Contains("\u65E5\u672C\u30AA\u30E9\u30AF\u30EB \u30CB\u30DB\u30F3"),
                "should not be able to find appended phonetic run");
            ex.Close();

        }
    }
}


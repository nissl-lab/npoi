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
using NUnit.Framework;
using NPOI.HSSF.Extractor;
using TestCases.HSSF;
using System.Text.RegularExpressions;
using NPOI.XSSF.Extractor;
using NPOI.XSSF;
using NPOI;

namespace TestCases.XSSF.Extractor
{

    /**
     * Tests for {@link XSSFExcelExtractor}
     */
    [TestFixture]
    public class TestXSSFExcelExtractor
    {
        protected XSSFExcelExtractor GetExtractor(String sampleName)
        {
            return new XSSFExcelExtractor(XSSFTestDataSamples.OpenSampleWorkbook(sampleName));
        }

        /**
         * Get text out of the simple file
         */
           [Test]
        public void TestGetSimpleText()
        {
            // a very simple file
            XSSFExcelExtractor extractor = GetExtractor("sample.xlsx");

            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);

            // Check sheet names
            Assert.IsTrue(text.StartsWith("Sheet1"));
            Assert.IsTrue(text.EndsWith("Sheet3\n"));

            // Now without, will have text
            extractor.SetIncludeSheetNames(false);
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
                "The quick brown fox jumps over the lazy dog\n\t" +
                "hello, xssf		hello, xssf\n\t" +
                "hello, xssf		hello, xssf\n\t" +
                "hello, xssf		hello, xssf\n\t" +
                "hello, xssf		hello, xssf\n";
            Assert.AreEqual(
                    CHUNK1 +
                    "at\t4995\n" +
                    CHUNK2
                    , text);

            // Now Get formulas not their values
            extractor.SetFormulasNotResults(true);
            text = extractor.Text;
            Assert.AreEqual(
                    CHUNK1 +
                    "at\tSUM(B1:B9)\n" +
                    CHUNK2, text);

            // With sheet names too
            extractor.SetIncludeSheetNames(true);
            text = extractor.Text;
            Assert.AreEqual(
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
            XSSFExcelExtractor extractor = GetExtractor("AverageTaxRates.xlsx");

            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);

            // Might not have all formatting it should do!
            // TODO decide if we should really have the "null" in there
            Assert.IsTrue(text.StartsWith(
                            "Avgtxfull\n" +
                            "\t\t(iii) AVERAGE TAX RATES ON ANNUAL"
            ));
            extractor.Close();
        }

        /**
         * Test that we return pretty much the same as
         *  ExcelExtractor does, when we're both passed
         *  the same file, just saved as xls and xlsx
         */
           [Test]
        public void TestComparedToOLE2()
        {
            // A fairly simple file - ooxml
            XSSFExcelExtractor ooxmlExtractor = GetExtractor("SampleSS.xlsx");

            ExcelExtractor ole2Extractor =
                new ExcelExtractor(HSSFTestDataSamples.OpenSampleWorkbook("SampleSS.xls"));

            POITextExtractor[] extractors =
                new POITextExtractor[] { ooxmlExtractor, ole2Extractor };
            for (int i = 0; i < extractors.Length; i++)
            {
                POITextExtractor extractor = extractors[i];

                String text = Regex.Replace(extractor.Text,"[\r\t]", "");
                Assert.IsTrue(text.StartsWith("First Sheet\nTest spreadsheet\n2nd row2nd row 2nd column\n"));
                Regex pattern = new Regex(".*13(\\.0+)?\\s+Sheet3.*",RegexOptions.Compiled);
                Assert.IsTrue(pattern.IsMatch(text));
                
            }
            ole2Extractor.Close();
            ooxmlExtractor.Close();
        }

        /**
         * From bug #45540
         */
           [Test]
           public void TestHeaderFooter()
           {
               String[] files = new String[] {
        "45540_classic_Header.xlsx", "45540_form_Header.xlsx",
        "45540_classic_Footer.xlsx", "45540_form_Footer.xlsx",
        };
               foreach (String sampleName in files)
               {
                   XSSFExcelExtractor extractor = GetExtractor(sampleName);
                   String text = extractor.Text;

                   Assert.IsTrue(text.Contains("testdoc"), "Unable to find expected word in text from " + sampleName + "\n" + text);
                   Assert.IsTrue(text.Contains("test phrase"), "Unable to find expected word in text\n" + text);
                   extractor.Close();
               }
           }

        /**
         * From bug #45544
         */
           [Test]
        public void TestComments()
        {

            XSSFExcelExtractor extractor = GetExtractor("45544.xlsx");
            String text = extractor.Text;

            // No comments there yet
            Assert.IsFalse(text.Contains("testdoc"), "Unable to find expected word in text\n" + text);
            Assert.IsFalse(text.Contains("test phrase"), "Unable to find expected word in text\n" + text);

            // Turn on comment extraction, will then be
            extractor.SetIncludeCellComments(true);
            text = extractor.Text;
            Assert.IsTrue(text.Contains("testdoc"), "Unable to find expected word in text\n" + text);
            Assert.IsTrue(text.Contains("test phrase"), "Unable to find expected word in text\n" + text);
            extractor.Close();
        }
        [Test]
        public void TestInlineStrings()
        {
            XSSFExcelExtractor extractor = GetExtractor("InlineStrings.xlsx");
            extractor.SetFormulasNotResults(true);
            String text = extractor.Text;

            // Numbers
            Assert.IsTrue(text.Contains("43"), "Unable to find expected word in text\n" + text);
            Assert.IsTrue(text.Contains("22"), "Unable to find expected word in text\n" + text);

            // Strings
            Assert.IsTrue(text.Contains("ABCDE"), "Unable to find expected word in text\n" + text);
            Assert.IsTrue(text.Contains("Long Text"), "Unable to find expected word in text\n" + text);

            // Inline Strings
            Assert.IsTrue(text.Contains("1st Inline String"), "Unable to find expected word in text\n" + text);
            Assert.IsTrue(text.Contains("And More"), "Unable to find expected word in text\n" + text);

            // Formulas
            Assert.IsTrue(text.Contains("A2"), "Unable to find expected word in text\n" + text);
            Assert.IsTrue(text.Contains("A5-A$2"), "Unable to find expected word in text\n" + text);
            extractor.Close();
        }

        [Test]
        public void TestEmptyCells()
        {
            XSSFExcelExtractor extractor = GetExtractor("SimpleNormal.xlsx");

            String text = extractor.Text;
            Assert.IsTrue(text.Length > 0);
            
            // This sheet demonstrates the preservation of empty cells, as
            // signified by sequential \t characters.
            Assert.AreEqual(
                // Sheet 1
                "Sheet1\n" + 
                "test\t\t1\n" + 
                "test 2\t\t2\n" + 
                "\t\t3\n" + 
                "\t\t4\n" + 
                "\t\t5\n" + 
                "\t\t6\n" + 
                // Sheet 2
                "Sheet Number 2\n" + 
                "This is sheet 2\n" + 
                "Stuff\n" + 
                "1\t2\t3\t4\t5\t6\n" + 
                "1/1/90\n" + 
                "10\t\t3\n", 
                text);

            extractor.Close();
        }

        /**
	     * Simple test for text box text
	     * @throws IOException
	     */
        [Test]
        public void TestTextBoxes()
        {
            XSSFExcelExtractor extractor = GetExtractor("WithTextBox.xlsx");
            try
            {
                extractor.SetFormulasNotResults(true);
                String text = extractor.Text;
                Assert.IsTrue(text.IndexOf("Line 1") > -1);
                Assert.IsTrue(text.IndexOf("Line 2") > -1);
                Assert.IsTrue(text.IndexOf("Line 3") > -1);
            }
            finally
            {
                extractor.Close();
            }
        }
    }
}



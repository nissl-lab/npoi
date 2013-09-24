/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
namespace TestCases.HWPF.Converter
{
    using System;
    using System.Xml;
    
    using NPOI.HWPF;
    using NPOI.HWPF.Converter;
    using NPOI.HWPF.UserModel;
    using NUnit.Framework;
    /**
     * Test cases for {@link WordToFoConverter}
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    //[TestFixture]
    public class TestWordToFoConverter
    {
        private static void assertContains(String result, String substring)
        {
            if (!result.Contains(substring))
                Assert.Fail("Substring \"" + substring
                        + "\" not found in the following string: \"" + result
                        + "\"");
        }

        private static String getFoText(String sampleFileName)
        {
            HWPFDocument hwpfDocument = new HWPFDocument(POIDataSamples.GetDocumentInstance().OpenResourceAsStream(sampleFileName));

            WordToFoConverter wordToFoConverter = new WordToFoConverter(new XmlDocument());
            wordToFoConverter.ProcessDocument(hwpfDocument);

            return wordToFoConverter.Document.InnerXml;
        }
        [Test]
        public void testDocumentProperties()
        {
            String result = getFoText("documentProperties.doc");

            assertContains(
                    result,
                    "<dc:title xmlns:dc=\"http://purl.org/dc/elements/1.1/\">This is document title</dc:title>");
            assertContains(
                    result,
                    "<pdf:Keywords xmlns:pdf=\"http://ns.adobe.com/pdf/1.3/\">This is document keywords</pdf:Keywords>");
        }
        [Test]
        public void testEndnote()
        {
            String result = getFoText("endingnote.doc");

            assertContains(result,
                    "<fo:basic-link id=\"endnote_back_1\" internal-destination=\"endnote_1\">");
            assertContains(result,
                    "<fo:inline baseline-shift=\"super\" font-size=\"smaller\">1</fo:inline>");
            assertContains(result,
                    "<fo:basic-link id=\"endnote_1\" internal-destination=\"endnote_back_1\">");
            assertContains(result,
                    "<fo:inline baseline-shift=\"super\" font-size=\"smaller\">1 </fo:inline>");
            assertContains(result, "Ending note text");
        }
        [Test]
        public void testEquation()
        {
            String sampleFileName = "equation.doc";
            String result = getFoText(sampleFileName);

            assertContains(result, "<!--Image link to '0.emf' can be here-->");
        }
        [Test]
        public void testHyperlink()
        {
            String sampleFileName = "hyperlink.doc";
            String result = getFoText(sampleFileName);

            assertContains(result,
                    "<fo:basic-link external-destination=\"http://testuri.org/\">");
            assertContains(result, "Hyperlink text");
        }
        [Test]
        public void testInnerTable()
        {
            String sampleFileName = "innertable.doc";
            String result = getFoText(sampleFileName);

            assertContains(result,
                    "padding-end=\"0.0in\" padding-start=\"0.0in\" width=\"1.0770833in\"");
        }
        [Test]
        public void testPageBreak()
        {
            String sampleFileName = "page-break.doc";
            String result = getFoText(sampleFileName);

            assertContains(result, "<fo:block break-before=\"page\"");
        }
        [Test]
        public void testPageBreakBefore()
        {
            String sampleFileName = "page-break-before.doc";
            String result = getFoText(sampleFileName);

            assertContains(result, "<fo:block break-before=\"page\"");
        }
        [Test]
        public void testPageref()
        {
            String sampleFileName = "pageref.doc";
            String result = getFoText(sampleFileName);

            assertContains(result,
                    "<fo:basic-link internal-destination=\"bookmark_userref\">");
            assertContains(result, "1");
            assertContains(result, "<fo:inline id=\"bookmark_userref\">");
        }
    }
}
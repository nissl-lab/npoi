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
namespace TestCases.HWPF.Converter
{
    using System;
    using System.Xml;
    
    using NPOI.HWPF;
    using NPOI.HWPF.Converter;
    using TestCases;
    using NUnit.Framework;

    /**
     * Test cases for {@link WordToHtmlConverter}
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    [TestFixture]
    public class TestWordToHtmlConverter
    {
        private static void AssertContains(String result, String substring)
        {
            if (!result.Contains(substring))
                Assert.Fail("Substring \"" + substring
                        + "\" not found in the following string: \"" + result
                        + "\"");
        }

        private static String GetHtmlText(String sampleFileName)
        {
            return GetHtmlText(sampleFileName, false);
        }
        
        private static String GetHtmlText(String sampleFileName,
                bool emulatePictureStorage)
        {
            HWPFDocument hwpfDocument = new HWPFDocument(POIDataSamples
                    .GetDocumentInstance().OpenResourceAsStream(sampleFileName));
            XmlDocument newDocument = new XmlDocument();
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                    newDocument);

            if (emulatePictureStorage)
            {
                //wordToHtmlConverter.SetPicturesManager( new PicturesManager()
                //{
                //    public String SavePicture( byte[] content,
                //            PictureType pictureType, String suggestedName )
                //    {
                //        return suggestedName;
                //    }
                //} );
            }

            wordToHtmlConverter.ProcessDocument(hwpfDocument);

            ;
            return wordToHtmlConverter.Document.InnerXml;
        }
        //[Test]
        public void TestAIOOBTap()
        {
            String result = GetHtmlText("AIOOB-Tap.doc");
            AssertContains(result.Substring(0, 6000), "<table class=\"t1\">");
        }
        //[Test]
        public void TestBug33519()
        {
            String result = GetHtmlText("Bug33519.doc");
            AssertContains(
                    result,
                    "\u041F\u043B\u0430\u043D\u0438\u043D\u0441\u043A\u0438 \u0442\u0443\u0440\u043E\u0432\u0435");
            AssertContains(result,
                    "\u042F\u0432\u043E\u0440 \u0410\u0441\u0435\u043D\u043E\u0432");
        }
        //[Test]
        public void TestBug46610_2()
        {
            String result = GetHtmlText("Bug46610_2.doc");
            AssertContains(
                    result,
                    "012345678911234567892123456789312345678941234567890123456789112345678921234567893123456789412345678");
        }
        //[Test]
        public void TestBug46817()
        {
            String result = GetHtmlText("Bug46817.doc");
            String substring = "<table class=\"t1\">";
            AssertContains(result, substring);
        }
        //[Test]
        public void TestBug47286()
        {
            String result = GetHtmlText("Bug47286.doc");

            Assert.IsFalse(result.Contains("FORMTEXT"));

            AssertContains(result, "color:#4f6228;");
            AssertContains(result, "Passport No and the date of expire");
            AssertContains(result, "mfa.gov.cy");
        }
        //[Test]
        public void TestBug48075()
        {
            GetHtmlText("Bug48075.doc");
        }
        [Test]
        public void TestDocumentProperties()
        {
            String result = GetHtmlText("documentProperties.doc");

            AssertContains(result, "<title>This is document title</title>");
            AssertContains(result,
                    "<meta name=\"keywords\" content=\"This is document keywords\" />");
        }
        //[Test]
        public void TestEmailhyperlink()
        {
            String result = GetHtmlText("Bug47286.doc");
            String substring = "provisastpet@mfa.gov.cy";
            AssertContains(result, substring);
        }
        //[Test]
        public void TestEndnote()
        {
            String result = GetHtmlText("endingnote.doc");

            AssertContains(
                    result,
                    "<a class=\"a1 endnoteanchor\" href=\"#endnote_1\" name=\"endnote_back_1\">1</a>");
            AssertContains(
                    result,
                    "<a class=\"a1 endnoteindex\" href=\"#endnote_back_1\" name=\"endnote_1\">1</a> <span");
            AssertContains(result, "Ending note text");
        }
        //[Test]
        public void TestEquation()
        {
            String result = GetHtmlText("equation.doc");

            AssertContains(result, "<!--Image link to '0.emf' can be here-->");
        }
        //[Test]
        public void TestHyperlink()
        {
            String result = GetHtmlText("hyperlink.doc");

            AssertContains(result, "<span>Before text; </span><a ");
            AssertContains(result,
                    "<a href=\"http://testuri.org/\"><span>Hyperlink text</span></a>");
            AssertContains(result, "</a><span>; after text</span>");
        }
        //[Test]
        public void TestInnerTable()
        {
            GetHtmlText("innertable.doc");
        }
        //[Test]
        public void TestO_kurs_doc()
        {
            GetHtmlText("o_kurs.doc");
        }
        //[Test]
        public void TestPageref()
        {
            String result = GetHtmlText("pageref.doc");

            AssertContains(result, "<a href=\"#userref\">");
            AssertContains(result, "<a name=\"userref\">");
            AssertContains(result, "1");
        }
        //[Test]
        public void TestPicture()
        {
            String result = GetHtmlText("picture.doc", true);

            // picture
            AssertContains(result, "src=\"0.emf\"");
            // visible size
            AssertContains(result, "width:3.1305554in;height:1.7250001in;");
            // shift due to crop
            AssertContains(result, "left:-0.09375;top:-0.25694445;");
            // size without crop
            AssertContains(result, "width:3.4125in;height:2.325in;");
        }
        //[Test]
        public void TestPicturesEscher()
        {
            String result = GetHtmlText("pictures_escher.doc", true);
            AssertContains(result, "<img src=\"s0.PNG\">");
            AssertContains(result, "<img src=\"s808.PNG\">");
        }
        //[Test]
        public void TestTableMerges()
        {
            String result = GetHtmlText("table-merges.doc");

            AssertContains(result, "<td class=\"td1\" colspan=\"3\">");
            AssertContains(result, "<td class=\"td2\" colspan=\"2\">");
        }
    }
}

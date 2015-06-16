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
using System.Collections.Generic;
using NUnit.Framework;
namespace NPOI.XWPF.UserModel
{
    [TestFixture]
    public class TestXWPFSDT
    {

        /**
         * Test simple tag and title extraction from SDT
         * @throws Exception
         */
        [Test]
        public void TestTagTitle()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Bug54849.docx");
            String tag = null;
            String title = null;
            List<XWPFSDT> sdts = extractAllSDTs(doc);
            foreach (XWPFSDT sdt in sdts)
            {
                if (sdt.Content.ToString().Equals("Rich_text"))
                {
                    tag = "MyTag";
                    title = "MyTitle";
                    break;
                }
            }
            Assert.AreEqual(12, sdts.Count, "controls size");

            Assert.AreEqual("MyTag", tag, "tag");
            Assert.AreEqual("MyTitle", title, "title");
        }

        [Test]
        public void TestGetSDTs()
        {
            String[] contents = new String[]{
                "header_rich_text",
                "Rich_text",
                "Rich_text_pre_table\nRich_text_cell1\t\t\t\n\nRich_text_post_table",
                "Plain_text_no_newlines",
                "Plain_text_with_newlines1\nplain_text_with_newlines2",
                "Watermelon",
                "Dirt",
                "4/16/2013",
                "rich_text_in_paragraph_in_cell",
                "Footer_rich_text",
                "Footnote_sdt",
                "Endnote_sdt"

        };
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Bug54849.docx");
            List<XWPFSDT> sdts = extractAllSDTs(doc);

            Assert.AreEqual(contents.Length, sdts.Count, "number of sdts");

            for (int i = 0; i < sdts.Count; i++)
            {//contents.Length; i++){
                XWPFSDT sdt = sdts[i];

                Assert.AreEqual(contents[i], sdt.Content.ToString(), i + ": " + contents[i]);
            }
        }
        [Test]
        public void TestFailureToGetSDTAsCell()
        {
            /**
             * The current code fails to extract an sdt if it comprises/is the parent
             * of a cell in a table.
             */
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Bug54849.docx");
            List<XWPFSDT> sdts = extractAllSDTs(doc);
            bool found = false;
            foreach (XWPFSDT sdt in sdts)
            {
                if (sdt.Content.Text.ToLower().IndexOf("rich_text_in_cell") > -1)
                {
                    found = true;
                }
            }
            Assert.AreEqual(false, found, "SDT as cell known failure");
        }
        /**
         * POI-55142 and Tika 1130
         */
        [Test]
        public void TestNewLinesBetweenRuns()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("Bug55142.docx");
            List<XWPFSDT> sdts = extractAllSDTs(doc);
            List<String> targs = new List<String>();
            //these test newlines and tabs in paragraphs/body elements
            targs.Add("Rich-text1 abcdefghi");
            targs.Add("Rich-text2 abcd\t\tefgh");
            targs.Add("Rich-text3 abcd\nefg");
            targs.Add("Rich-text4 abcdefg");
            targs.Add("Rich-text5 abcdefg\nhijk");
            targs.Add("Plain-text1 abcdefg");
            targs.Add("Plain-text2 abcdefg\nhijk\nlmnop");
            //this tests consecutive runs within a cell (not a paragraph)
            //this test case was triggered by Tika-1130
            targs.Add("sdt_incell2 abcdefg");

            for (int i = 0; i < sdts.Count; i++)
            {
                XWPFSDT sdt = sdts[i];
                Assert.AreEqual(targs[i], sdt.Content.Text, targs[i]);
            }
        }
        private List<XWPFSDT> extractAllSDTs(XWPFDocument doc)
        {

            List<XWPFSDT> sdts = new List<XWPFSDT>();

            IList<XWPFHeader> headers = doc.HeaderList;
            foreach (XWPFHeader header in headers)
            {
                sdts.AddRange(ExtractSDTsFromBodyElements(header.BodyElements));
            }
            sdts.AddRange(ExtractSDTsFromBodyElements(doc.BodyElements));

            IList<XWPFFooter> footers = doc.FooterList;
            foreach (XWPFFooter footer in footers)
            {
                sdts.AddRange(ExtractSDTsFromBodyElements(footer.BodyElements));
            }

            foreach (XWPFFootnote footnote in doc.GetFootnotes())
            {
                sdts.AddRange(ExtractSDTsFromBodyElements(footnote.BodyElements));
            }
            foreach (KeyValuePair<int, XWPFFootnote> e in doc.Endnotes)
            {
                sdts.AddRange(ExtractSDTsFromBodyElements(e.Value.BodyElements));
            }
            return sdts;
        }

        private List<XWPFSDT> ExtractSDTsFromBodyElements(IList<IBodyElement> elements)
        {
            List<XWPFSDT> sdts = new List<XWPFSDT>();
            foreach (IBodyElement e in elements)
            {
                if (e is XWPFSDT)
                {
                    XWPFSDT sdt = (XWPFSDT)e;
                    sdts.Add(sdt);
                }
                else if (e is XWPFParagraph)
                {

                    XWPFParagraph p = (XWPFParagraph)e;
                    foreach (IRunElement e2 in p.IRuns)
                    {
                        if (e2 is XWPFSDT)
                        {
                            XWPFSDT sdt = (XWPFSDT)e2;
                            sdts.Add(sdt);
                        }
                    }
                }
                else if (e is XWPFTable)
                {
                    XWPFTable table = (XWPFTable)e;
                    sdts.AddRange(extractSDTsFromTable(table));
                }
            }
            return sdts;
        }

        private List<XWPFSDT> extractSDTsFromTable(XWPFTable table)
        {
            List<XWPFSDT> sdts = new List<XWPFSDT>();
            foreach (XWPFTableRow r in table.Rows)
            {
                foreach (XWPFTableCell c in r.GetTableCells())
                {
                    sdts.AddRange(ExtractSDTsFromBodyElements(c.BodyElements));
                }
            }
            return sdts;
        }
    }

}
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
namespace NPOI.XWPF.Extractor
{
    using System;
    using NPOI.XWPF.Model;
    using NPOI.XWPF.UserModel;
    using NPOI.OpenXml4Net.OPC;
    using System.Text;
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.IO;
    using System.Xml;

    /**
     * Helper class to extract text from an OOXML Word file
     */
    public class XWPFWordExtractor : POIXMLTextExtractor
    {
        public static XWPFRelation[] SUPPORTED_TYPES = new XWPFRelation[] {
      XWPFRelation.DOCUMENT, XWPFRelation.TEMPLATE,
      XWPFRelation.MACRO_DOCUMENT, 
      XWPFRelation.MACRO_TEMPLATE_DOCUMENT
   };

        private XWPFDocument document;
        private bool fetchHyperlinks = false;

        public XWPFWordExtractor(OPCPackage Container)
            : this(new XWPFDocument(Container))
        {

        }
        public XWPFWordExtractor(XWPFDocument document)
            : base(document)
        {

            this.document = document;
        }

        /**
         * Should we also fetch the hyperlinks, when fetching 
         *  the text content? Default is to only output the
         *  hyperlink label, and not the contents
         */
        public void SetFetchHyperlinks(bool fetch)
        {
            fetchHyperlinks = fetch;
        }


        public override String Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                XWPFHeaderFooterPolicy hfPolicy = document.GetHeaderFooterPolicy();

                // Start out with all headers
                ExtractHeaders(text, hfPolicy);

                // body elements
                foreach (IBodyElement e in document.BodyElements)
                {
                    AppendBodyElementText(text, e);
                    text.Append('\n');
                }

                // Finish up with all the footers
                ExtractFooters(text, hfPolicy);

                return text.ToString();
            }
        }

        public void AppendBodyElementText(StringBuilder text, IBodyElement e)
        {
            if (e is XWPFParagraph paragraph)
            {
                AppendParagraphText(text, paragraph);
            }
            else if (e is XWPFTable table)
            {
                AppendTableText(text, table);
            }
            else if (e is XWPFSDT xwpfsdt)
            {
                text.Append(xwpfsdt.Content.Text);
            }
        }

        public void AppendParagraphText(StringBuilder text, XWPFParagraph paragraph)
        {
            try
            {
                CT_SectPr ctSectPr = null;
                if (paragraph.GetCTP().pPr != null)
                {
                    ctSectPr = paragraph.GetCTP().pPr.sectPr;
                }

                XWPFHeaderFooterPolicy headerFooterPolicy = null;

                if (ctSectPr != null)
                {
                    headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
                    ExtractHeaders(text, headerFooterPolicy);
                }


                foreach (IRunElement run in paragraph.Runs)
                {
                    text.Append(run.ToString());
                    if (run is XWPFHyperlinkRun hyperlinkRun && fetchHyperlinks)
                    {
                        XWPFHyperlink link = hyperlinkRun.GetHyperlink(document);
                        if (link != null)
                            text.Append(" <" + link.URL + ">");
                    }
                }

                // Add comments
                XWPFCommentsDecorator decorator = new XWPFCommentsDecorator(paragraph, null);
                String commentText = decorator.GetCommentText();
                if (commentText.Length > 0)
                {
                    text.Append(commentText).Append('\n');
                }

                // Do endnotes and footnotes
                String footnameText = paragraph.FootnoteText;
                if (footnameText != null && footnameText.Length > 0)
                {
                    text.Append(footnameText + '\n');
                }

                if (ctSectPr != null)
                {
                    ExtractFooters(text, headerFooterPolicy);
                }
            }
            catch (IOException e)
            {
                throw new POIXMLException(e);
            }
            catch (XmlException e)
            {
                throw new POIXMLException(e);
            }

        }

        private void AppendTableText(StringBuilder text, XWPFTable table)
        {
            //this works recursively to pull embedded tables from tables
            foreach (XWPFTableRow row in table.Rows)
            {
                List<ICell> cells = row.GetTableICells();
                for (int i = 0; i < cells.Count; i++)
                {
                    ICell cell = cells[(i)];
                    if (cell is XWPFTableCell tableCell)
                    {
                        text.Append(tableCell.GetTextRecursively());
                    }
                    else if (cell is XWPFSDTCell xwpfsdtCell)
                    {
                        text.Append(xwpfsdtCell.Content.Text);
                    }
                    if (i < cells.Count - 1)
                    {
                        text.Append("\t");
                    }
                }
                text.Append('\n');
            }
        }

        private void ExtractFooters(StringBuilder text, XWPFHeaderFooterPolicy hfPolicy)
        {
            if (hfPolicy == null) return;
            if (hfPolicy.GetFirstPageFooter() != null)
            {
                text.Append(hfPolicy.GetFirstPageFooter().Text);
            }
            if (hfPolicy.GetEvenPageFooter() != null)
            {
                text.Append(hfPolicy.GetEvenPageFooter().Text);
            }
            if (hfPolicy.GetDefaultFooter() != null)
            {
                text.Append(hfPolicy.GetDefaultFooter().Text);
            }
        }

        private void ExtractHeaders(StringBuilder text, XWPFHeaderFooterPolicy hfPolicy)
        {
            if (hfPolicy == null) return;
            if (hfPolicy.GetFirstPageHeader() != null)
            {
                text.Append(hfPolicy.GetFirstPageHeader().Text);
            }
            if (hfPolicy.GetEvenPageHeader() != null)
            {
                text.Append(hfPolicy.GetEvenPageHeader().Text);
            }
            if (hfPolicy.GetDefaultHeader() != null)
            {
                text.Append(hfPolicy.GetDefaultHeader().Text);
            }
        }
    }

}
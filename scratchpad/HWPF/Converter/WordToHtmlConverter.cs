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
using System.Collections.Generic;
using System.Text;
using NPOI.Util;
using NPOI.HWPF.UserModel;
using System.Xml;

namespace NPOI.HWPF.Converter
{
    /**
     * Converts Word files (95-2007) into HTML files.
     * <p>
     * This implementation doesn't create images or links to them. This can be
     * changed by overriding {@link #processImage(XmlElement, boolean, Picture)}
     * method.
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    /// <summary>
    /// Beta
    /// </summary>
    public class WordToHtmlConverter : AbstractWordConverter
    {
        /**
         * Holds properties values, applied to current <tt>p</tt> element. Those
         * properties shall not be doubled in children <tt>span</tt> elements.
         */
        private class BlockProperies
        {
            public String pFontName;
            public int pFontSize;

            public BlockProperies(String pFontName, int pFontSize)
            {
                this.pFontName = pFontName;
                this.pFontSize = pFontSize;
            }
        }

        private static POILogger logger = POILogFactory.GetLogger(typeof(WordToHtmlConverter));

        private static String GetSectionStyle(Section section)
        {
            float leftMargin = section.MarginLeft / AbstractWordUtils.TWIPS_PER_INCH;
            float rightMargin = section.MarginRight / AbstractWordUtils.TWIPS_PER_INCH;
            float topMargin = section.MarginTop / AbstractWordUtils.TWIPS_PER_INCH;
            float bottomMargin = section.MarginBottom / AbstractWordUtils.TWIPS_PER_INCH;

            String style = "margin: " + topMargin + "in " + rightMargin + "in "
                    + bottomMargin + "in " + leftMargin + "in;";

            if (section.NumColumns > 1)
            {
                style += "column-count: " + (section.NumColumns) + ";";
                if (section.IsColumnsEvenlySpaced)
                {
                    float distance = section.DistanceBetweenColumns / AbstractWordUtils.TWIPS_PER_INCH;
                    style += "column-gap: " + distance + "in;";
                }
                else
                {
                    style += "column-gap: 0.25in;";
                }
            }
            return style;
        }

        public static XmlDocument Process(string docFile)
        {
            HWPFDocumentCore wordDocument = WordToHtmlUtils.LoadDoc(docFile);
            XmlDocument xmlDoc = new XmlDocument();
            //WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(xmlDoc);
            // wordToHtmlConverter.ProcessDocument(wordDocument);
            return xmlDoc;
        }

        private Stack<BlockProperies> blocksProperies = new Stack<BlockProperies>();

        private HtmlDocumentFacade htmlDocumentFacade;

        private XmlElement notes = null;

        /**
         * Creates new instance of {@link WordToHtmlConverter}. Can be used for
         * output several {@link HWPFDocument}s into single HTML document.
         * 
         * @param document
         *            XML DOM Document used as HTML document
         */
        public WordToHtmlConverter(XmlDocument document)
        {
            this.htmlDocumentFacade = new HtmlDocumentFacade(document);
        }

        protected override void AfterProcess()
        {
            if (notes != null)
                htmlDocumentFacade.Body.AppendChild(notes);

            htmlDocumentFacade.UpdateStylesheet();
        }

        public override XmlDocument Document
        {
            get
            {
                return htmlDocumentFacade.Document;
            }
        }

        protected override void OutputCharacters(XmlElement pElement, CharacterRun characterRun, string text)
        {
            XmlElement span = htmlDocumentFacade.Document.CreateElement("span");
            pElement.AppendChild(span);

            StringBuilder style = new StringBuilder();
            BlockProperies blockProperies = this.blocksProperies.Peek();
            Triplet triplet = GetCharacterRunTriplet(characterRun);

            if (!string.IsNullOrEmpty(triplet.fontName)
                    && !WordToHtmlUtils.Equals(triplet.fontName,
                            blockProperies.pFontName))
            {
                style.Append("font-family:" + triplet.fontName + ";");
            }
            if (characterRun.GetFontSize() / 2 != blockProperies.pFontSize)
            {
                style.Append("font-size:" + characterRun.GetFontSize() / 2 + "pt;");
            }
            if (triplet.bold)
            {
                style.Append("font-weight:bold;");
            }
            if (triplet.italic)
            {
                style.Append("font-style:italic;");
            }

            WordToHtmlUtils.AddCharactersProperties(characterRun, style);
            if (style.Length != 0)
                htmlDocumentFacade.AddStyleClass(span, "s", style.ToString());

            XmlText textNode = htmlDocumentFacade.CreateText(text);
            span.AppendChild(textNode);
        }

        protected override void ProcessBookmarks(HWPFDocumentCore wordDocument, XmlElement currentBlock, Range range, int currentTableLevel, IList<Bookmark> rangeBookmarks)
        {
            XmlElement parent = currentBlock;
            foreach (Bookmark bookmark in rangeBookmarks)
            {
                XmlElement bookmarkElement = htmlDocumentFacade.CreateBookmark(bookmark.Name);
                parent.AppendChild(bookmarkElement);
                parent = bookmarkElement;
            }

            if (range != null)
                ProcessCharacters(wordDocument, currentTableLevel, range, parent);
            throw new NotImplementedException();
        }

        protected override void ProcessDocumentInformation(NPOI.HPSF.SummaryInformation summaryInformation)
        {
            if (!string.IsNullOrEmpty(summaryInformation.Title))
                htmlDocumentFacade.Title = summaryInformation.Title;

            if (!string.IsNullOrEmpty(summaryInformation.Author))
                htmlDocumentFacade.AddAuthor(summaryInformation.Author);

            if (!string.IsNullOrEmpty(summaryInformation.Keywords))
                htmlDocumentFacade.AddKeywords(summaryInformation.Keywords);

            if (!string.IsNullOrEmpty(summaryInformation.Comments))
                htmlDocumentFacade.AddDescription(summaryInformation.Comments);
        }
        public void processDocumentPart(HWPFDocumentCore wordDocument, Range range)
        {
            base.ProcessDocumentPart(wordDocument, range);
            AfterProcess();
        }
        protected override void ProcessDrawnObject(HWPFDocument doc, CharacterRun characterRun, OfficeDrawing officeDrawing, string path, XmlElement block)
        {
            XmlElement img = htmlDocumentFacade.CreateImage(path);
            block.AppendChild(img);
        }

        protected override void ProcessEndnoteAutonumbered(HWPFDocument wordDocument, int noteIndex, XmlElement block, Range endnoteTextRange)
        {
            ProcessNoteAutonumbered(wordDocument, "end", noteIndex, block, endnoteTextRange);
        }

        private void ProcessNoteAutonumbered(HWPFDocument wordDocument, string type, int noteIndex, XmlElement block, Range noteTextRange)
        {
            String textIndex = (noteIndex + 1).ToString();
            String textIndexClass = htmlDocumentFacade.GetOrCreateCssClass("a", "a", "vertical-align:super;font-size:smaller;");
            String forwardNoteLink = type + "note_" + textIndex;
            String backwardNoteLink = type + "note_back_" + textIndex;

            XmlElement anchor = htmlDocumentFacade.CreateHyperlink("#" + forwardNoteLink);
            anchor.SetAttribute("name", backwardNoteLink);
            anchor.SetAttribute("class", textIndexClass + " " + type + "noteanchor");
            anchor.InnerText = textIndex;
            block.AppendChild(anchor);

            if (notes == null)
            {
                notes = htmlDocumentFacade.CreateBlock();
                notes.SetAttribute("class", "notes");
            }

            XmlElement note = htmlDocumentFacade.CreateBlock();
            note.SetAttribute("class", type + "note");
            notes.AppendChild(note);

            XmlElement bookmark = htmlDocumentFacade.CreateBookmark(forwardNoteLink);
            bookmark.SetAttribute("href", "#" + backwardNoteLink);
            bookmark.InnerText = (textIndex);
            bookmark.SetAttribute("class", textIndexClass + " " + type  + "noteindex");
            note.AppendChild(bookmark);
            note.AppendChild(htmlDocumentFacade.CreateText(" "));

            XmlElement span = htmlDocumentFacade.Document.CreateElement("span");
            span.SetAttribute("class", type + "notetext");
            note.AppendChild(span);

            this.blocksProperies.Push(new BlockProperies("", -1));
            try
            {
                ProcessCharacters(wordDocument, int.MinValue, noteTextRange, span);
            }
            finally
            {
                this.blocksProperies.Pop();
            }
        }

        protected override void ProcessFootnoteAutonumbered(HWPFDocument wordDocument, int noteIndex, XmlElement block, Range footnoteTextRange)
        {
            ProcessNoteAutonumbered(wordDocument, "foot", noteIndex, block, footnoteTextRange);
        }

        protected override void ProcessHyperlink(HWPFDocumentCore wordDocument, XmlElement currentBlock, Range textRange, int currentTableLevel, string hyperlink)
        {
            XmlElement basicLink = htmlDocumentFacade.CreateHyperlink(hyperlink);
            currentBlock.AppendChild(basicLink);

            if (textRange != null)
                ProcessCharacters(wordDocument, currentTableLevel, textRange, basicLink);
        }
        /**
         * This method shall store image bytes in external file and convert it if
         * necessary. Images shall be stored using PNG format. Other formats may be
         * not supported by user browser.
         * <p>
         * Please note the {@link #processImage(XmlElement, boolean, Picture, String)}.
         * 
         * @param currentBlock
         *            currently processed HTML element, like <tt>p</tt>. Shall be
         *            used as parent of newly created <tt>img</tt>
         * @param inlined
         *            if image is inlined
         * @param picture
         *            HWPF object, contained picture data and properties
         */
        protected override void ProcessImage(XmlElement currentBlock, bool inlined, Picture picture)
        {
            PicturesManager fileManager = GetPicturesManager();
            if (fileManager != null)
            {
                String url = fileManager
                        .SavePicture(picture.GetContent(),
                                picture.SuggestPictureType(),
                                picture.SuggestFullFileName());

                if (!string.IsNullOrEmpty(url))
                {
                    ProcessImage(currentBlock, inlined, picture, url);
                    return;
                }
            }

            // no default implementation -- skip
            currentBlock.AppendChild(htmlDocumentFacade.Document
                    .CreateComment("Image link to '"
                            + picture.SuggestFullFileName() + "' can be here"));
        }
        protected void ProcessImage(XmlElement currentBlock, bool inlined,
            Picture picture, String imageSourcePath)
        {
            int aspectRatioX = picture.HorizontalScalingFactor;
            int aspectRatioY = picture.VerticalScalingFactor;

            StringBuilder style = new StringBuilder();

            float imageWidth;
            float imageHeight;

            float cropTop;
            float cropBottom;
            float cropLeft;
            float cropRight;

            if (aspectRatioX > 0)
            {
                imageWidth = picture.DxaGoal * aspectRatioX / 1000
                        / AbstractWordUtils.TWIPS_PER_INCH;
                cropRight = picture.DxaCropRight * aspectRatioX / 1000
                        / AbstractWordUtils.TWIPS_PER_INCH;
                cropLeft = picture.DxaCropLeft * aspectRatioX / 1000
                        / AbstractWordUtils.TWIPS_PER_INCH;
            }
            else
            {
                imageWidth = picture.DxaGoal / AbstractWordUtils.TWIPS_PER_INCH;
                cropRight = picture.DxaCropRight / AbstractWordUtils.TWIPS_PER_INCH;
                cropLeft = picture.DxaCropLeft / AbstractWordUtils.TWIPS_PER_INCH;
            }

            if (aspectRatioY > 0)
            {
                imageHeight = picture.DyaGoal * aspectRatioY / 1000
                        / AbstractWordUtils.TWIPS_PER_INCH;
                cropTop = picture.DyaCropTop * aspectRatioY / 1000
                        / AbstractWordUtils.TWIPS_PER_INCH;
                cropBottom = picture.DyaCropBottom * aspectRatioY / 1000
                        / AbstractWordUtils.TWIPS_PER_INCH;
            }
            else
            {
                imageHeight = picture.DyaGoal / AbstractWordUtils.TWIPS_PER_INCH;
                cropTop = picture.DyaCropTop / AbstractWordUtils.TWIPS_PER_INCH;
                cropBottom = picture.DyaCropBottom / AbstractWordUtils.TWIPS_PER_INCH;
            }

            XmlElement root;
            if (cropTop != 0 || cropRight != 0 || cropBottom != 0 || cropLeft != 0)
            {
                float visibleWidth = Math.Max(0, imageWidth - cropLeft - cropRight);
                float visibleHeight = Math.Max(0, imageHeight - cropTop - cropBottom);

                root = htmlDocumentFacade.CreateBlock();
                htmlDocumentFacade.AddStyleClass(root, "d", "vertical-align:text-bottom;width:" + visibleWidth + "in;height:" + visibleHeight + "in;");

                // complex
                XmlElement inner = htmlDocumentFacade.CreateBlock();
                htmlDocumentFacade.AddStyleClass(inner, "d", "position:relative;width:" + visibleWidth + "in;height:" + visibleHeight + "in;overflow:hidden;");
                root.AppendChild(inner);

                XmlElement image = htmlDocumentFacade.CreateImage(imageSourcePath);
                htmlDocumentFacade.AddStyleClass(image, "i", "position:absolute;left:-" + cropLeft + ";top:-" + cropTop + ";width:" + imageWidth + "in;height:" + imageHeight + "in;");
                inner.AppendChild(image);

                style.Append("overflow:hidden;");
            }
            else
            {
                root = htmlDocumentFacade.CreateImage(imageSourcePath);
                root.SetAttribute("style", "width:" + imageWidth + "in;height:" + imageHeight + "in;vertical-align:text-bottom;");
            }

            currentBlock.AppendChild(root);
        }
        protected override void ProcessLineBreak(XmlElement block, CharacterRun characterRun)
        {
            block.AppendChild(htmlDocumentFacade.CreateLineBreak());
        }

        protected override void ProcessPageBreak(HWPFDocumentCore wordDocument, XmlElement flow)
        {
            flow.AppendChild(htmlDocumentFacade.CreateLineBreak());
        }

        protected override void ProcessPageref(HWPFDocumentCore wordDocument, XmlElement currentBlock, Range textRange, int currentTableLevel, string pageref)
        {
            XmlElement basicLink = htmlDocumentFacade.CreateHyperlink("#" + pageref);
            currentBlock.AppendChild(basicLink);

            if (textRange != null)
                ProcessCharacters(wordDocument, currentTableLevel, textRange, basicLink);
        }

        protected override void ProcessParagraph(HWPFDocumentCore wordDocument, XmlElement parentElement, int currentTableLevel, Paragraph paragraph, string bulletText)
        {
            XmlElement pElement = htmlDocumentFacade.CreateParagraph();
            parentElement.AppendChild(pElement);

            StringBuilder style = new StringBuilder();
            WordToHtmlUtils.AddParagraphProperties(paragraph, style);

            int charRuns = paragraph.NumCharacterRuns;

            if (charRuns == 0)
            {
                return;
            }

            {
                String pFontName;
                int pFontSize;
                CharacterRun characterRun = paragraph.GetCharacterRun(0);
                if (characterRun != null)
                {
                    Triplet triplet = GetCharacterRunTriplet(characterRun);
                    pFontSize = characterRun.GetFontSize() / 2;
                    pFontName = triplet.fontName;
                    WordToHtmlUtils.AddFontFamily(pFontName, style);
                    WordToHtmlUtils.AddFontSize(pFontSize, style);
                }
                else
                {
                    pFontSize = -1;
                    pFontName = string.Empty;
                }
                blocksProperies.Push(new BlockProperies(pFontName, pFontSize));
            }
            try
            {
                if (!string.IsNullOrEmpty(bulletText))
                {
                    XmlText textNode = htmlDocumentFacade.CreateText(bulletText);
                    pElement.AppendChild(textNode);
                }

                ProcessCharacters(wordDocument, currentTableLevel, paragraph, pElement);
            }
            finally
            {
                blocksProperies.Pop();
            }

            if (style.Length > 0)
                htmlDocumentFacade.AddStyleClass(pElement, "p", style.ToString());

            WordToHtmlUtils.CompactSpans(pElement);
        }

        protected override void ProcessSection(HWPFDocumentCore wordDocument, Section section, int s)
        {
            XmlElement div = htmlDocumentFacade.CreateBlock();
            htmlDocumentFacade.AddStyleClass(div, "d", GetSectionStyle(section));
            htmlDocumentFacade.Body.AppendChild(div);

            ProcessParagraphes(wordDocument, div, section, int.MinValue);
        }

        protected void processSingleSection(HWPFDocumentCore wordDocument,
            Section section)
        {
            htmlDocumentFacade.AddStyleClass(htmlDocumentFacade.Body, "b", GetSectionStyle(section));

            ProcessParagraphes(wordDocument, htmlDocumentFacade.Body, section, int.MinValue);
        }

        protected override void ProcessTable(HWPFDocumentCore wordDocument, XmlElement flow, Table table)
        {
            XmlElement tableHeader = htmlDocumentFacade.CreateTableHeader();
            XmlElement tableBody = htmlDocumentFacade.CreateTableBody();

            int[] tableCellEdges = WordToHtmlUtils.BuildTableCellEdgesArray(table);
            int tableRows = table.NumRows;

            int maxColumns = int.MinValue;
            for (int r = 0; r < tableRows; r++)
            {
                maxColumns = Math.Max(maxColumns, table.GetRow(r).NumCells());
            }

            for (int r = 0; r < tableRows; r++)
            {
                TableRow tableRow = table.GetRow(r);

                XmlElement tableRowElement = htmlDocumentFacade.CreateTableRow();
                StringBuilder tableRowStyle = new StringBuilder();
                WordToHtmlUtils.AddTableRowProperties(tableRow, tableRowStyle);

                // index of current element in tableCellEdges[]
                int currentEdgeIndex = 0;
                int rowCells = tableRow.NumCells();
                for (int c = 0; c < rowCells; c++)
                {
                    TableCell tableCell = tableRow.GetCell(c);

                    if (tableCell.IsVerticallyMerged() && !tableCell.IsFirstVerticallyMerged())
                    {
                        currentEdgeIndex += getTableCellEdgesIndexSkipCount(table, r, tableCellEdges, currentEdgeIndex, c, tableCell);
                        continue;
                    }

                    XmlElement tableCellElement;
                    if (tableRow.isTableHeader())
                    {
                        tableCellElement = htmlDocumentFacade.CreateTableHeaderCell();
                    }
                    else
                    {
                        tableCellElement = htmlDocumentFacade.CreateTableCell();
                    }
                    StringBuilder tableCellStyle = new StringBuilder();
                    WordToHtmlUtils.AddTableCellProperties(tableRow, tableCell, r == 0, r == tableRows - 1, c == 0, c == rowCells - 1, tableCellStyle);

                    int colSpan = GetNumberColumnsSpanned(tableCellEdges, currentEdgeIndex, tableCell);
                    currentEdgeIndex += colSpan;

                    if (colSpan == 0)
                        continue;

                    if (colSpan != 1)
                        tableCellElement.SetAttribute("colspan", colSpan.ToString());

                    int rowSpan = GetNumberRowsSpanned(table, r, c,
                           tableCell);
                    if (rowSpan > 1)
                        tableCellElement.SetAttribute("rowspan", rowSpan.ToString());

                    ProcessParagraphes(wordDocument, tableCellElement, tableCell, 0  /*table.TableLevel Todo: */);

                    if (!tableCellElement.HasChildNodes)
                    {
                        tableCellElement.AppendChild(htmlDocumentFacade.CreateParagraph());
                    }
                    if (tableCellStyle.Length > 0)
                        htmlDocumentFacade.AddStyleClass(tableCellElement, tableCellElement.LocalName, tableCellStyle.ToString());

                    tableRowElement.AppendChild(tableCellElement);
                }

                if (tableRowStyle.Length > 0)
                    tableRowElement.SetAttribute("class", htmlDocumentFacade.GetOrCreateCssClass("tr", "r", tableRowStyle.ToString()));

                if (tableRow.isTableHeader())
                {
                    tableHeader.AppendChild(tableRowElement);
                }
                else
                {
                    tableBody.AppendChild(tableRowElement);
                }
            }

            XmlElement tableElement = htmlDocumentFacade.CreateTable();
            tableElement.SetAttribute("class",
                            htmlDocumentFacade.GetOrCreateCssClass(tableElement.LocalName, "t", "table-layout:fixed;border-collapse:collapse;border-spacing:0;"));
            if (tableHeader.HasChildNodes)
            {
                tableElement.AppendChild(tableHeader);
            }
            if (tableBody.HasChildNodes)
            {
                tableElement.AppendChild(tableBody);
                flow.AppendChild(tableElement);
            }
            else
            {
                logger.Log(POILogger.WARN, "Table without body starting at [", table.StartOffset.ToString(), "; ", table.EndOffset.ToString(), ")");
            }
        }
    }
}

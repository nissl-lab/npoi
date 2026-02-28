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
namespace NPOI.HWPF.Converter
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using NPOI.HPSF;
    using NPOI.HWPF;
    using NPOI.HWPF.UserModel;
    using NPOI.Util;


    /**
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */

    public class WordToFoConverter : AbstractWordConverter
    {

        private static POILogger logger = POILogFactory.GetLogger(typeof(WordToFoConverter));


        public static XmlDocument Process(string docFile)
        {
            HWPFDocumentCore hwpfDocument = WordToFoUtils.LoadDoc(docFile);
            WordToFoConverter wordToFoConverter = new WordToFoConverter(new XmlDocument());
            wordToFoConverter.ProcessDocument(hwpfDocument);
            return wordToFoConverter.Document;
        }

        private List<XmlElement> endnotes = new List<XmlElement>(0);

        protected FoDocumentFacade foDocumentFacade;

        private object objLinkCounter = new object();
        private int internalLinkCounter = 0;

        private bool outputCharactersLanguage = false;

        private LinkedList<String> usedIds = new LinkedList<String>();

        /**
         * Creates new instance of {@link WordToFoConverter}. Can be used for output
         * several {@link HWPFDocument}s into single FO document.
         * 
         * @param document
         *            XML DOM Document used as XSL FO document. Shall support
         *            namespaces
         */
        public WordToFoConverter(XmlDocument document)
        {
            this.foDocumentFacade = new FoDocumentFacade(document);
        }

        protected XmlElement CreateNoteInline(String noteIndexText)
        {
            XmlElement inline = foDocumentFacade.CreateInline();
            inline.InnerText = (noteIndexText);
            inline.SetAttribute("baseline-shift", "super");
            inline.SetAttribute("font-size", "smaller");
            return inline;
        }

        protected String CreatePageMaster(NPOI.HWPF.UserModel.Section section, String type, int sectionIndex)
        {
            float height = section.PageHeight / WordToFoUtils.TWIPS_PER_INCH;
            float width = section.PageWidth / WordToFoUtils.TWIPS_PER_INCH;
            float leftMargin = section.MarginLeft
                    / WordToFoUtils.TWIPS_PER_INCH;
            float rightMargin = section.MarginRight
                    / WordToFoUtils.TWIPS_PER_INCH;
            float topMargin = section.MarginTop / WordToFoUtils.TWIPS_PER_INCH;
            float bottomMargin = section.MarginBottom
                    / WordToFoUtils.TWIPS_PER_INCH;

            // add these to the header
            String pageMasterName = type + "-page" + sectionIndex;

            XmlElement pageMaster = foDocumentFacade.AddSimplePageMaster(pageMasterName);
            pageMaster.SetAttribute("page-height", height + "in");
            pageMaster.SetAttribute("page-width", width + "in");

            XmlElement regionBody = foDocumentFacade.AddRegionBody(pageMaster);
            regionBody.SetAttribute("margin", topMargin + "in " + rightMargin
                    + "in " + bottomMargin + "in " + leftMargin + "in");

            /*
             * 6.4.14 fo:region-body
             * 
             * The values of the padding and border-width traits must be "0".
             */
            // WordToFoUtils.setBorder(regionBody, sep.getBrcTop(), "top");
            // WordToFoUtils.setBorder(regionBody, sep.getBrcBottom(), "bottom");
            // WordToFoUtils.setBorder(regionBody, sep.getBrcLeft(), "left");
            // WordToFoUtils.setBorder(regionBody, sep.getBrcRight(), "right");

            if (section.NumColumns > 1)
            {
                regionBody.SetAttribute("column-count", "" + (section.NumColumns));
                if (section.IsColumnsEvenlySpaced)
                {
                    float distance = section.DistanceBetweenColumns / WordToFoUtils.TWIPS_PER_INCH;
                    regionBody.SetAttribute("column-gap", distance + "in");
                }
                else
                {
                    regionBody.SetAttribute("column-gap", "0.25in");
                }
            }

            return pageMasterName;
        }

        public override XmlDocument Document
        {
            get
            {
                return foDocumentFacade.Document;
            }
        }

        public bool IsOutputCharactersLanguage()
        {
            return outputCharactersLanguage;
        }

        protected override void OutputCharacters(XmlElement block, CharacterRun characterRun,
                String text)
        {
            XmlElement inline = foDocumentFacade.CreateInline();

            Triplet triplet = GetCharacterRunTriplet(characterRun);

            if (!string.IsNullOrEmpty(triplet.fontName))
                WordToFoUtils.SetFontFamily(inline, triplet.fontName);
            WordToFoUtils.SetBold(inline, triplet.bold);
            WordToFoUtils.SetItalic(inline, triplet.italic);
            WordToFoUtils.SetFontSize(inline, characterRun.GetFontSize() / 2);
            WordToFoUtils.SetCharactersProperties(characterRun, inline);

            if (IsOutputCharactersLanguage())
                WordToFoUtils.SetLanguage(characterRun, inline);

            block.AppendChild(inline);

            XmlText textNode = foDocumentFacade.CreateText(text);
            inline.AppendChild(textNode);
        }

        protected override void ProcessBookmarks(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range range, int currentTableLevel,
                IList<Bookmark> rangeBookmarks)
        {
            XmlElement parent = currentBlock;
            foreach (Bookmark bookmark in rangeBookmarks)
            {
                XmlElement bookmarkElement = foDocumentFacade.CreateInline();
                String idName = "bookmark_" + bookmark.Name;
                // make sure ID used once
                if (SetId(bookmarkElement, idName))
                {
                    /*
                     * if it just empty fo:inline without "id" attribute doesn't
                     * making sense to add it to DOM
                     */
                    parent.AppendChild(bookmarkElement);
                    parent = bookmarkElement;
                }
            }

            if (range != null)
                ProcessCharacters(wordDocument, currentTableLevel, range, parent);
        }

        protected override void ProcessDocumentInformation(SummaryInformation summaryInformation)
        {
            if (!string.IsNullOrEmpty(summaryInformation.Title))
                foDocumentFacade.SetTitle(summaryInformation.Title);

            if (!string.IsNullOrEmpty(summaryInformation.Author))
                foDocumentFacade.SetCreator(summaryInformation.Author);

            if (!string.IsNullOrEmpty(summaryInformation.Keywords))
                foDocumentFacade.SetKeywords(summaryInformation.Keywords);

            if (!string.IsNullOrEmpty(summaryInformation.Comments))
                foDocumentFacade.SetDescription(summaryInformation.Comments);
        }

        protected override void ProcessDrawnObject(HWPFDocument doc,
                CharacterRun characterRun, OfficeDrawing officeDrawing,
                String path, XmlElement block)
        {
            XmlElement externalGraphic = foDocumentFacade.CreateExternalGraphic(path);
            block.AppendChild(externalGraphic);
        }

        protected override void ProcessEndnoteAutonumbered(HWPFDocument wordDocument,
                int noteIndex, XmlElement block, Range endnoteTextRange)
        {
            String textIndex;// = (internalLinkCounter.incrementAndGet()).ToString();
            lock (objLinkCounter)
            {
                internalLinkCounter++;

                textIndex = internalLinkCounter.ToString();
            }
            String forwardLinkName = "endnote_" + textIndex;
            String backwardLinkName = "endnote_back_" + textIndex;

            XmlElement forwardLink = foDocumentFacade
                    .CreateBasicLinkInternal(forwardLinkName);
            forwardLink.AppendChild(CreateNoteInline(textIndex));
            SetId(forwardLink, backwardLinkName);
            block.AppendChild(forwardLink);

            XmlElement endnote = foDocumentFacade.CreateBlock();
            XmlElement backwardLink = foDocumentFacade
                    .CreateBasicLinkInternal(backwardLinkName);
            backwardLink.AppendChild(CreateNoteInline(textIndex + " "));
            SetId(backwardLink, forwardLinkName);
            endnote.AppendChild(backwardLink);

            ProcessCharacters(wordDocument, int.MinValue, endnoteTextRange, endnote);

            WordToFoUtils.CompactInlines(endnote);
            this.endnotes.Add(endnote);
        }

        protected override void ProcessFootnoteAutonumbered(HWPFDocument wordDocument,
                int noteIndex, XmlElement block, Range footnoteTextRange)
        {
            String textIndex;// = (internalLinkCounter.incrementAndGet()).ToString();
            lock (objLinkCounter)
            {
                internalLinkCounter++;

                textIndex = internalLinkCounter.ToString();
            }
            String forwardLinkName = "footnote_" + textIndex;
            String backwardLinkName = "footnote_back_" + textIndex;

            XmlElement footNote = foDocumentFacade.CreateFootnote();
            block.AppendChild(footNote);

            XmlElement inline = foDocumentFacade.CreateInline();
            XmlElement forwardLink = foDocumentFacade
                    .CreateBasicLinkInternal(forwardLinkName);
            forwardLink.AppendChild(CreateNoteInline(textIndex));
            SetId(forwardLink, backwardLinkName);
            inline.AppendChild(forwardLink);
            footNote.AppendChild(inline);

            XmlElement footnoteBody = foDocumentFacade.CreateFootnoteBody();
            XmlElement footnoteBlock = foDocumentFacade.CreateBlock();
            XmlElement backwardLink = foDocumentFacade
                    .CreateBasicLinkInternal(backwardLinkName);
            backwardLink.AppendChild(CreateNoteInline(textIndex + " "));
            SetId(backwardLink, forwardLinkName);
            footnoteBlock.AppendChild(backwardLink);
            footnoteBody.AppendChild(footnoteBlock);
            footNote.AppendChild(footnoteBody);

            ProcessCharacters(wordDocument, int.MinValue, footnoteTextRange, footnoteBlock);

            WordToFoUtils.CompactInlines(footnoteBlock);
        }

        protected override void ProcessHyperlink(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range textRange, int currentTableLevel,
                String hyperlink)
        {
            XmlElement basicLink = foDocumentFacade
                    .CreateBasicLinkExternal(hyperlink);
            currentBlock.AppendChild(basicLink);

            if (textRange != null)
                ProcessCharacters(wordDocument, currentTableLevel, textRange, basicLink);
        }

        /**
         * This method shall store image bytes in external file and convert it if
         * necessary. Images shall be stored using PNG format (for bitmap) or SVG
         * (for vector). Other formats may be not supported by your XSL FO
         * processor.
         * <p>
         * Please note the
         * {@link WordToFoUtils#setPictureProperties(Picture, XmlElement)} method.
         * 
         * @param currentBlock
         *            currently processed FO element, like <tt>fo:block</tt>. Shall
         *            be used as parent of newly created
         *            <tt>fo:external-graphic</tt> or
         *            <tt>fo:instream-foreign-object</tt>
         * @param inlined
         *            if image is inlined
         * @param picture
         *            HWPF object, contained picture data and properties
         */
        protected override void ProcessImage(XmlElement currentBlock, bool inlined,
                Picture picture)
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
            currentBlock.AppendChild(foDocumentFacade.Document
                    .CreateComment("Image link to '"
                            + picture.SuggestFullFileName() + "' can be here"));
        }

        protected void ProcessImage(XmlElement currentBlock, bool inlined,
                Picture picture, String url)
        {
            XmlElement externalGraphic = foDocumentFacade
                   .CreateExternalGraphic(url);
            WordToFoUtils.SetPictureProperties(picture, externalGraphic);
            currentBlock.AppendChild(externalGraphic);
        }

        protected override void ProcessLineBreak(XmlElement block, CharacterRun characterRun)
        {
            block.AppendChild(foDocumentFacade.CreateBlock());
        }

        protected override void ProcessPageBreak(HWPFDocumentCore wordDocument, XmlElement flow)
        {
            XmlElement block = null;
            XmlNodeList childNodes = flow.ChildNodes;
            if (childNodes.Count > 0)
            {
                XmlNode lastChild = childNodes[childNodes.Count - 1];
                if (lastChild is XmlElement)
                {
                    XmlElement lastElement = (XmlElement)lastChild;
                    if (!lastElement.HasAttribute("break-after"))
                    {
                        block = lastElement;
                    }
                }
            }

            if (block == null)
            {
                block = foDocumentFacade.CreateBlock();
                flow.AppendChild(block);
            }
            block.SetAttribute("break-after", "page");
        }

        protected override void ProcessPageref(HWPFDocumentCore hwpfDocument,
                XmlElement currentBlock, Range textRange, int currentTableLevel,
                String pageref)
        {
            XmlElement basicLink = foDocumentFacade
                    .CreateBasicLinkInternal("bookmark_" + pageref);
            currentBlock.AppendChild(basicLink);

            if (textRange != null)
                ProcessCharacters(hwpfDocument, currentTableLevel, textRange,
                        basicLink);
        }

        protected override void ProcessParagraph(HWPFDocumentCore hwpfDocument,
                XmlElement parentFopElement, int currentTableLevel,
                Paragraph paragraph, String bulletText)
        {
            XmlElement block = foDocumentFacade.CreateBlock();
            parentFopElement.AppendChild(block);

            WordToFoUtils.SetParagraphProperties(paragraph, block);

            int charRuns = paragraph.NumCharacterRuns;

            if (charRuns == 0)
            {
                return;
            }

            bool haveAnyText = false;

            if (!string.IsNullOrEmpty(bulletText))
            {
                XmlElement inline = foDocumentFacade.CreateInline();
                block.AppendChild(inline);

                XmlText textNode = foDocumentFacade.CreateText(bulletText);
                inline.AppendChild(textNode);

                haveAnyText |= bulletText.Trim().Length != 0;
            }

            haveAnyText = ProcessCharacters(hwpfDocument, currentTableLevel, paragraph, block);

            if (!haveAnyText)
            {
                XmlElement leader = foDocumentFacade.CreateLeader();
                block.AppendChild(leader);
            }

            WordToFoUtils.CompactInlines(block);
            return;
        }

        protected override void ProcessSection(HWPFDocumentCore wordDocument,
                NPOI.HWPF.UserModel.Section section, int sectionCounter)
        {
            String regularPage = CreatePageMaster(section, "page", sectionCounter);

            XmlElement pageSequence = foDocumentFacade.AddPageSequence(regularPage);
            XmlElement flow = foDocumentFacade.AddFlowToPageSequence(pageSequence, "xsl-region-body");

            ProcessParagraphes(wordDocument, flow, section, int.MinValue);

            if (endnotes != null && endnotes.Count != 0)
            {
                foreach (XmlElement endnote in endnotes)
                    flow.AppendChild(endnote);
                endnotes.Clear();
            }
        }

        protected override void ProcessTable(HWPFDocumentCore wordDocument, XmlElement flow,
                Table table)
        {
            XmlElement tableHeader = foDocumentFacade.CreateTableHeader();
            XmlElement tableBody = foDocumentFacade.CreateTableBody();

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

                XmlElement tableRowElement = foDocumentFacade.CreateTableRow();
                WordToFoUtils.SetTableRowProperties(tableRow, tableRowElement);

                // index of current element in tableCellEdges[]
                int currentEdgeIndex = 0;
                int rowCells = tableRow.NumCells();
                for (int c = 0; c < rowCells; c++)
                {
                    TableCell tableCell = tableRow.GetCell(c);

                    if (tableCell.IsVerticallyMerged() && !tableCell.IsFirstVerticallyMerged())
                    {
                        currentEdgeIndex += getTableCellEdgesIndexSkipCount(table,
                                r, tableCellEdges, currentEdgeIndex, c, tableCell);
                        continue;
                    }

                    XmlElement tableCellElement = foDocumentFacade.CreateTableCell();
                    WordToFoUtils.SetTableCellProperties(tableRow, tableCell,
                            tableCellElement, r == 0, r == tableRows - 1, c == 0,
                            c == rowCells - 1);

                    int colSpan = GetNumberColumnsSpanned(tableCellEdges,
                            currentEdgeIndex, tableCell);
                    currentEdgeIndex += colSpan;

                    if (colSpan == 0)
                        continue;

                    if (colSpan != 1)
                        tableCellElement.SetAttribute("number-columns-spanned", (colSpan).ToString());

                    int rowSpan = GetNumberRowsSpanned(table, r, c, tableCell);
                    if (rowSpan > 1)
                        tableCellElement.SetAttribute("number-rows-spanned", (rowSpan).ToString());

                    ProcessParagraphes(wordDocument, tableCellElement, tableCell,
                            table.TableLevel);

                    if (!tableCellElement.HasChildNodes)
                    {
                        tableCellElement.AppendChild(foDocumentFacade
                                .CreateBlock());
                    }

                    tableRowElement.AppendChild(tableCellElement);
                }

                if (tableRowElement.HasChildNodes)
                {
                    if (tableRow.isTableHeader())
                    {
                        tableHeader.AppendChild(tableRowElement);
                    }
                    else
                    {
                        tableBody.AppendChild(tableRowElement);
                    }
                }
            }

            XmlElement tableElement = foDocumentFacade.CreateTable();
            tableElement.SetAttribute("table-layout", "fixed");
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
                logger.Log(POILogger.WARN, "Table without body starting on offset " + table.StartOffset + " -- " + table.EndOffset);
            }
        }

        protected bool SetId(XmlElement element, String id)
        {
            // making sure ID used once
            if (usedIds.Contains(id))
            {
                logger.Log(POILogger.WARN,
                        "Tried to create element with same ID '", id, "'. Skipped");
                return false;
            }

            element.SetAttribute("id", id);
            usedIds.AddLast(id);
            return true;
        }

        public void SetOutputCharactersLanguage(bool outputCharactersLanguage)
        {
            this.outputCharactersLanguage = outputCharactersLanguage;
        }

    }
}
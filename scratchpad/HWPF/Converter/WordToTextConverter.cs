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
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using System.Reflection;



    public class WordToTextConverter : AbstractWordConverter
    {
        private static POILogger logger = POILogFactory
                .GetLogger(typeof(WordToTextConverter));

        public static String GetText( DirectoryNode root )  
        {
             HWPFDocumentCore wordDocument = AbstractWordUtils.LoadDoc( root );
            return GetText( wordDocument );
        }

        public static String GetText(string docFile)
        {
            HWPFDocumentCore wordDocument = AbstractWordUtils
                   .LoadDoc(docFile);
            return GetText(wordDocument);
        }

        public static String GetText(HWPFDocumentCore wordDocument)
        {
            WordToTextConverter wordToTextConverter = new WordToTextConverter(new XmlDocument());
            wordToTextConverter.ProcessDocument(wordDocument);
            return wordToTextConverter.GetText();
        }



        public static XmlDocument Process(string docFile)
        {
            HWPFDocumentCore wordDocument = AbstractWordUtils
                   .LoadDoc(docFile);
            WordToTextConverter wordToTextConverter = new WordToTextConverter(new XmlDocument());
            wordToTextConverter.ProcessDocument(wordDocument);
            return wordToTextConverter.Document;
        }

        //private AtomicInteger noteCounters = new AtomicInteger(1);
        private int noteCounters = 1;
        private object objCounters = new object();

        private XmlElement notes = null;

        private bool outputSummaryInformation = false;

        private TextDocumentFacade textDocumentFacade;

        /**
         * Creates new instance of {@link WordToTextConverter}. Can be used for
         * output several {@link HWPFDocument}s into single text document.
         * 
         * @throws ParserConfigurationException
         *             if an internal {@link DocumentBuilder} cannot be created
         */
        public WordToTextConverter()
        {
            this.textDocumentFacade = new TextDocumentFacade(new XmlDocument());
        }

        /**
         * Creates new instance of {@link WordToTextConverter}. Can be used for
         * output several {@link HWPFDocument}s into single text document.
         * 
         * @param document
         *            XML DOM XmlDocument used as storage for text pieces
         */
        public WordToTextConverter(XmlDocument document)
        {
            this.textDocumentFacade = new TextDocumentFacade(document);
        }

        protected override void AfterProcess()
        {
            if (notes != null)
                textDocumentFacade.Body.AppendChild(notes);
        }

        public override XmlDocument Document
        {
            get
            {
                return textDocumentFacade.Document;
            }
        }

        public String GetText()
        {
            //StringWriter stringWriter = new StringWriter();
            //DOMSource domSource = new DOMSource( GetDocument() );
            //StreamResult streamResult = new StreamResult( stringWriter );

            //TransformerFactory tf = TransformerFactory.NewInstance();
            //Transformer serializer = tf.NewTransformer();
            //// TODO set encoding from a command argument
            //serializer.SetOutputProperty( OutputKeys.ENCODING, "UTF-8" );
            //serializer.SetOutputProperty( OutputKeys.INDENT, "no" );
            //serializer.SetOutputProperty( OutputKeys.METHOD, "text" );
            //serializer.Transform( domSource, streamResult );

            //return stringWriter.ToString();
            return textDocumentFacade.Document.InnerText;
        }

        public bool IsOutputSummaryInformation()
        {
            return outputSummaryInformation;
        }

        protected override void OutputCharacters(XmlElement block, CharacterRun characterRun,
                String text)
        {
            block.AppendChild(textDocumentFacade.CreateText(text));
        }

        protected override void ProcessBookmarks(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range range, int currentTableLevel,
                IList<Bookmark> rangeBookmarks)
        {
            ProcessCharacters(wordDocument, currentTableLevel, range, currentBlock);
        }

        protected override void ProcessDocumentInformation(
                SummaryInformation summaryInformation)
        {
            if (IsOutputSummaryInformation())
            {
                if (!string.IsNullOrEmpty(summaryInformation.Title))
                    textDocumentFacade.Title = (summaryInformation.Title);

                if (!string.IsNullOrEmpty(summaryInformation.Author))
                    textDocumentFacade.AddAuthor(summaryInformation.Author);

                if (!string.IsNullOrEmpty(summaryInformation.Comments))
                    textDocumentFacade.AddDescription(summaryInformation.Comments);

                if (!string.IsNullOrEmpty(summaryInformation.Keywords))
                    textDocumentFacade.AddKeywords(summaryInformation.Keywords);
            }
        }

        protected override void ProcessDocumentPart(HWPFDocumentCore wordDocument,
                Range range)
        {
            base.ProcessDocumentPart(wordDocument, range);
            AfterProcess();
        }

        protected override void ProcessDrawnObject(HWPFDocument doc,
                CharacterRun characterRun, OfficeDrawing officeDrawing,
                String path, XmlElement block)
        {
            // ignore
        }

        protected override void ProcessEndnoteAutonumbered(HWPFDocument wordDocument,
                int noteIndex, XmlElement block, Range endnoteTextRange)
        {
            ProcessNote(wordDocument, block, endnoteTextRange);
        }
        protected override void ProcessFootnoteAutonumbered(HWPFDocument wordDocument,
                int noteIndex, XmlElement block, Range footnoteTextRange)
        {
            ProcessNote(wordDocument, block, footnoteTextRange);
        }

        protected override void ProcessHyperlink(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range textRange, int currentTableLevel,
                String hyperlink)
        {
            ProcessCharacters(wordDocument, currentTableLevel, textRange,
                    currentBlock);

            currentBlock.AppendChild(textDocumentFacade.CreateText(" ("
                    + UNICODECHAR_ZERO_WIDTH_SPACE
                    + hyperlink.Replace("\\/", UNICODECHAR_ZERO_WIDTH_SPACE
                            + "\\/" + UNICODECHAR_ZERO_WIDTH_SPACE)
                    + UNICODECHAR_ZERO_WIDTH_SPACE + ")"));
        }

        protected override void ProcessImage(XmlElement currentBlock, bool inlined,
                Picture picture)
        {
            // ignore
        }

        protected override void ProcessLineBreak(XmlElement block, CharacterRun characterRun)
        {
            block.AppendChild(textDocumentFacade.CreateText("\n"));
        }

        protected void ProcessNote(HWPFDocument wordDocument, XmlElement block,
                Range noteTextRange)
        {
            int noteIndex;
            lock (objCounters)
            {
                noteIndex = noteCounters++;
            }
            block.AppendChild(textDocumentFacade
                    .CreateText(UNICODECHAR_ZERO_WIDTH_SPACE + "[" + noteIndex
                            + "]" + UNICODECHAR_ZERO_WIDTH_SPACE));

            if (notes == null)
                notes = textDocumentFacade.CreateBlock();

            XmlElement note = textDocumentFacade.CreateBlock();
            notes.AppendChild(note);

            note.AppendChild(textDocumentFacade.CreateText("^" + noteIndex
                    + "\t "));
            ProcessCharacters(wordDocument, int.MinValue, noteTextRange, note);
            note.AppendChild(textDocumentFacade.CreateText("\n"));
        }

        protected override bool ProcessOle2(HWPFDocument wordDocument, XmlElement block, Entry entry)
        {
            if (!(entry is DirectoryNode))
                return false;
            DirectoryNode directoryNode = (DirectoryNode)entry;

            /*
             * even if there is no ExtractorFactory in classpath, still support
             * included Word's objects
             */

            //TODO: Not completed
            if ( directoryNode.HasEntry( "WordDocument" ) )
            {
                String text = WordToTextConverter.GetText( (DirectoryNode) entry );
                block.AppendChild( textDocumentFacade
                        .CreateText( UNICODECHAR_ZERO_WIDTH_SPACE + text
                                + UNICODECHAR_ZERO_WIDTH_SPACE ) );
                return true;
            }

            Object extractor;
            
            /*try
            {
                Class<?> cls = Class
                        .ForName( "org.apache.poi.extractor.ExtractorFactory" );
                Method createExtractor = cls.GetMethod( "createExtractor",
                        DirectoryNode.class );
                extractor = createExtractor.Invoke( null, directoryNode );
            }
            catch ( Error exc )
            {
                // no extractor in classpath
                logger.Log( POILogger.WARN, "There is an OLE object entry '",
                        entry.GetName(),
                        "', but there is no text extractor for this object type ",
                        "or text extractor factory is not available: ", "" + exc );
                return false;
            }

            try
            {
                Method getText = extractor.GetClass().GetMethod( "getText" );
                String text = (String) getText.Invoke( extractor );

                block.AppendChild( textDocumentFacade
                        .CreateText( UNICODECHAR_ZERO_WIDTH_SPACE + text
                                + UNICODECHAR_ZERO_WIDTH_SPACE ) );
                return true;
            }
            catch ( Exception exc )
            {
                logger.Log( POILogger.ERROR,
                        "Unable to extract text from OLE entry '", entry.GetName(),
                        "': ", exc, exc );
                return false;
            }
             * */
            return false;
        }

        protected override void ProcessPageBreak(HWPFDocumentCore wordDocument, XmlElement flow)
        {
            XmlElement block = textDocumentFacade.CreateBlock();
            block.AppendChild(textDocumentFacade.CreateText("\n"));
            flow.AppendChild(block);
        }

        protected override void ProcessPageref(HWPFDocumentCore wordDocument,
                XmlElement currentBlock, Range textRange, int currentTableLevel,
                String pageref)
        {
            ProcessCharacters(wordDocument, currentTableLevel, textRange,
                    currentBlock);
        }

        protected override void ProcessParagraph(HWPFDocumentCore wordDocument,
                XmlElement parentElement, int currentTableLevel, Paragraph paragraph,
                String bulletText)
        {
            XmlElement pElement = textDocumentFacade.CreateParagraph();
            pElement.AppendChild(textDocumentFacade.CreateText(bulletText));
            ProcessCharacters(wordDocument, currentTableLevel, paragraph, pElement);
            pElement.AppendChild(textDocumentFacade.CreateText("\n"));
            parentElement.AppendChild(pElement);
        }

        protected override void ProcessSection(HWPFDocumentCore wordDocument,
                NPOI.HWPF.UserModel.Section section, int s)
        {
            XmlElement sectionElement = textDocumentFacade.CreateBlock();
            ProcessParagraphes(wordDocument, sectionElement, section, int.MinValue);
            sectionElement.AppendChild(textDocumentFacade.CreateText("\n"));
            textDocumentFacade.Body.AppendChild(sectionElement);
        }

        protected override void ProcessTable(HWPFDocumentCore wordDocument, XmlElement flow,
                Table table)
        {
            int tableRows = table.NumRows;
            for (int r = 0; r < tableRows; r++)
            {
                TableRow tableRow = table.GetRow(r);

                XmlElement tableRowElement = textDocumentFacade.CreateTableRow();

                int rowCells = tableRow.NumCells();
                for (int c = 0; c < rowCells; c++)
                {
                    TableCell tableCell = tableRow.GetCell(c);

                    XmlElement tableCellElement = textDocumentFacade.CreateTableCell();

                    if (c != 0)
                        tableCellElement.AppendChild(textDocumentFacade
                                .CreateText("\t"));
                    ProcessCharacters(wordDocument, table.TableLevel, tableCell, tableCellElement);
                    tableRowElement.AppendChild(tableCellElement);
                }

                tableRowElement.AppendChild(textDocumentFacade.CreateText("\n"));
                flow.AppendChild(tableRowElement);
            }
        }

        public void SetOutputSummaryInformation(bool outputDocumentInformation)
        {
            this.outputSummaryInformation = outputDocumentInformation;
        }

    }
}

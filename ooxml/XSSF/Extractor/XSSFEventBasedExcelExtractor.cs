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

namespace NPOI.XSSF.Extractor
{

    using NPOI;
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.SS.Extractor;
    using NPOI.Util;
    using NPOI.XSSF.EventUserModel;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.UserModel;
    using System.Globalization;
    using NSAX;
    using NSAX.AElfred;
    using static NPOI.XSSF.EventUserModel.XSSFSheetXMLHandler;
    using NPOI.OpenXml4Net;

    /// <summary>
    /// Implementation of a text extractor from OOXML Excel
    ///  files that uses SAX event based parsing.
    /// </summary>
    public class XSSFEventBasedExcelExtractor : POIXMLTextExtractor, IExcelExtractor
    {

        //private static  POILogger LOGGER = POILogFactory.GetLogger(XSSFEventBasedExcelExtractor.class);

        private OPCPackage container;
        private POIXMLProperties properties;

        private CultureInfo locale;
        private bool includeTextBoxes = true;
        private bool includeSheetNames = true;
        private bool includeCellComments = false;
        private bool includeHeadersFooters = true;
        private bool formulasNotResults = false;
        private bool concatenatePhoneticRuns = true;

        public XSSFEventBasedExcelExtractor(String path)
            : this(OPCPackage.Open(path))
        {
        }
        public XSSFEventBasedExcelExtractor(OPCPackage container)
            : base(null)
        {


            this.container = container;

            properties = new POIXMLProperties(container);
        }

        public static void main(String[] args)
        {

            if(args.Length < 1)
            {
                Console.WriteLine("Use:");
                Console.WriteLine("  XSSFEventBasedExcelExtractor <filename.xlsx>");
                return;
            }
            var extractor = new XSSFEventBasedExcelExtractor(args[0]);
            Console.WriteLine(extractor.Text);
            extractor.Close();
        }

        /// <summary>
        /// Get or Set should sheet names be included? Default is true
        /// </summary>
        public bool IncludeSheetNames
        {
            get
            {
                return includeSheetNames;
            }
            set
            {
                includeSheetNames = value;
            }
        }


        /// <summary>
        /// Should we return the formula itself, and not
        /// the result it produces? Default is false
        /// </summary>
        public bool FormulasNotResults
        {
            get { return formulasNotResults; }
            set { formulasNotResults = value; }
        }

        /// <summary>
        /// Should headers and footers be included? Default is true
        /// </summary>
        public bool IncludeHeadersFooters
        {
            get { return includeHeadersFooters; }
            set { includeHeadersFooters = value; }
        }


        /// <summary>
        /// Should text from textboxes be included? Default is true
        /// </summary>
        public bool IncludeTextBoxes
        {
            get { return includeTextBoxes; }
            set { includeTextBoxes = value; }
        }

        /// <summary>
        /// </summary>
        /// <return>whether cell comments should be included</return>
        ///
        /// @since 3.16-beta3
        public bool IncludeCellComments
        {
            get { return includeCellComments; }
            set { this.includeCellComments = value; }
        }

        public bool AddTabEachEmptyCell { get; set; } = true;

        /// <summary>
        /// Concatenate text from &lt;rPh&gt; text elements in SharedStringsTable
        /// Default is true;
        /// </summary>
        /// <param name="concatenatePhoneticRuns">concatenatePhoneticRuns</param>
        public void SetConcatenatePhoneticRuns(bool concatenatePhoneticRuns)
        {
            this.concatenatePhoneticRuns = concatenatePhoneticRuns;
        }

        /// <summary>CultureInfo
        /// </summary>
        public CultureInfo Locale
        {
            get { return locale; }
            set { locale = value; }
        }
        /// <summary>
        /// Returns the opened OPCPackage container.
        /// </summary>
        public OPCPackage GetPackage()
        {
            return container;
        }

        /// <summary>
        /// Returns the core document properties
        /// </summary>
        public override CoreProperties GetCoreProperties()
        {
            return properties.CoreProperties;
        }
        /// <summary>
        /// Returns the extended document properties
        /// </summary>
        public override ExtendedProperties GetExtendedProperties()
        {
            return properties.ExtendedProperties;
        }
        /// <summary>
        /// Returns the custom document properties
        /// </summary>
        public override CustomProperties GetCustomProperties()
        {
            return properties.CustomProperties;
        }



        /// <summary>
        /// Processes the given sheet
        /// </summary>
        public void ProcessSheet(
                SheetContentsHandler sheetContentsExtractor,
                StylesTable styles,
                CommentsTable comments,
                ReadOnlySharedStringsTable strings,
                Stream sheetInputStream)

        {


            DataFormatter formatter;
            if(locale == null)
            {
                formatter = new DataFormatter();
            }
            else
            {
                formatter = new DataFormatter(locale);
            }

            InputSource sheetSource = new InputSource(sheetInputStream);
            try
            {
                SAXDriver sheetParser = new SAXDriver();
                IContentHandler handler = new XSSFSheetXMLHandler(
                styles, comments, strings, sheetContentsExtractor, formatter, formulasNotResults);
                sheetParser.ContentHandler = (handler);
                sheetParser.Parse(sheetSource);
            }
            catch(SAXException e)
            {
                throw new RuntimeException("SAX parser appears to be broken - " + e.Message);
            }
        }

        /// <summary>
        /// Processes the file and returns the text
        /// </summary>
        public override String Text
        {
            get
            {
                try
                {
                    ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(container, concatenatePhoneticRuns);
                    XSSFReader xssfReader = new XSSFReader(container);
                    StylesTable styles = xssfReader.StylesTable;
                    XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.GetSheetsData();

                    StringBuilder text = new StringBuilder();
                    SheetTextExtractor sheetExtractor = new SheetTextExtractor(this);

                    while(iter.MoveNext())
                    {
                        Stream stream = iter.Current;
                        if(includeSheetNames)
                        {
                            text.Append(iter.SheetName);
                            text.Append('\n');
                        }
                        CommentsTable comments = includeCellComments ? iter.SheetComments : null;
                        ProcessSheet(sheetExtractor, styles, comments, strings, stream);
                        if(includeHeadersFooters)
                        {
                            sheetExtractor.AppendHeaderText(text);
                        }
                        sheetExtractor.AppendCellText(text);
                        if(includeTextBoxes)
                        {
                            ProcessShapes(iter.Shapes, text);
                        }
                        if(includeHeadersFooters)
                        {
                            sheetExtractor.AppendFooterText(text);
                        }
                        sheetExtractor.Reset();
                        stream.Close();
                    }

                    return text.ToString();
                }
                catch(IOException)
                {
                    //LOGGER.log(POILogger.WARN, e);
                    return null;
                }
                catch(SAXException)
                {
                    //LOGGER.log(POILogger.WARN, se);
                    return null;
                }
                catch(OpenXml4NetException)
                {
                    //LOGGER.log(POILogger.WARN, o4je);
                    return null;
                }
            }
            
        }

        static void ProcessShapes(List<XSSFShape> shapes, StringBuilder text)
        {
            if(shapes == null)
            {
                return;
            }
            foreach(XSSFShape shape in shapes)
            {
                if(shape is XSSFSimpleShape)
                {
                    String sText = ((XSSFSimpleShape)shape).Text;
                    if(sText != null && sText.Length > 0)
                    {
                        text.Append(sText).Append('\n');
                    }
                }
            }
        }
        public override void Close()
        {

            if(container != null)
            {
                container.Close();
                container = null;
            }
            base.Close();
        }

        protected class SheetTextExtractor : SheetContentsHandler
        {
            private  StringBuilder output;
            private bool firstCellOfRow;
            private  Dictionary<String, String> headerFooterMap;
            private XSSFEventBasedExcelExtractor eb;
            public SheetTextExtractor(XSSFEventBasedExcelExtractor eb)
            {
                this.eb = eb;
                this.output = new StringBuilder();
                this.firstCellOfRow = true;
                this.headerFooterMap = eb.IncludeHeadersFooters ? new Dictionary<String, String>() : null;
            }
            public void StartRow(int rowNum)
            {
                firstCellOfRow = true;
            }
            public void EndRow(int rowNum)
            {
                output.Append('\n');
            }
            public void Cell(String cellRef, String formattedValue, XSSFComment comment)
            {
                if(firstCellOfRow)
                {
                    firstCellOfRow = false;
                }
                else
                {
                    output.Append('\t');
                }
                if(formattedValue != null)
                {
                    eb.CheckMaxTextSize(output, formattedValue);
                    output.Append(formattedValue);
                }
                if(eb.IncludeCellComments && comment != null)
                {
                    String commentText = comment.String.String.Replace('\n', ' ');
                    output.Append(formattedValue != null ? " Comment by " : "Comment by ");
                    eb.CheckMaxTextSize(output, commentText);
                    if(commentText.StartsWith(comment.Author + ": "))
                    {
                        output.Append(commentText);
                    }
                    else
                    {
                        output.Append(comment.Author).Append(": ").Append(commentText);
                    }
                }
            }
            public void HeaderFooter(String text, bool IsHeader, String tagName)
            {
                if(headerFooterMap != null)
                {
                    headerFooterMap[tagName] = text;
                }
            }

            /// <summary>
            /// Append the text for the named header or footer if found.
            /// </summary>
            private void AppendHeaderFooterText(StringBuilder buffer, String name)
            {
                String text = headerFooterMap.TryGetValue(name, out string value) ? value : null;
                if(text != null && text.Length > 0)
                {
                    // this is a naive way of handling the left, center, and right
                    // header and footer delimiters, but it seems to be as good as
                    // the method used by XSSFExcelExtractor
                    text = HandleHeaderFooterDelimiter(text, "&L");
                    text = HandleHeaderFooterDelimiter(text, "&C");
                    text = HandleHeaderFooterDelimiter(text, "&R");
                    buffer.Append(text).Append('\n');
                }
            }
            /// <summary>
            /// Remove the delimiter if its found at the beginning of the text,
            /// or replace it with a tab if its in the middle.
            /// </summary>
            private static String HandleHeaderFooterDelimiter(String text, String delimiter)
            {
                int index = text.IndexOf(delimiter);
                if(index == 0)
                {
                    text = text.Substring(2);
                }
                else if(index > 0)
                {
                    text = text.Substring(0, index) + "\t" + text.Substring(index + 2);
                }
                return text;
            }


            /// <summary>
            /// Append the text for each header type in the same order
            /// they are appended in XSSFExcelExtractor.
            /// </summary>
            /// <see cref="XSSFExcelExtractor.Text" />
            /// <see cref="NPOI.HSSF.extractor.ExcelExtractor._extractHeaderFooter(NPOI.SS.UserModel.HeaderFooter)" />
            public void AppendHeaderText(StringBuilder buffer)
            {
                AppendHeaderFooterText(buffer, "firstHeader");
                AppendHeaderFooterText(buffer, "oddHeader");
                AppendHeaderFooterText(buffer, "evenHeader");
            }

            /// <summary>
            /// Append the text for each footer type in the same order
            /// they are appended in XSSFExcelExtractor.
            /// </summary>
            /// <see cref="XSSFExcelExtractor.Text" />
            /// <see cref="NPOI.HSSF.extractor.ExcelExtractor._extractHeaderFooter(NPOI.SS.UserModel.HeaderFooter)" />
            public void AppendFooterText(StringBuilder buffer)
            {
                // append the text for each footer type in the same order
                // they are appended in XSSFExcelExtractor
                AppendHeaderFooterText(buffer, "firstFooter");
                AppendHeaderFooterText(buffer, "oddFooter");
                AppendHeaderFooterText(buffer, "evenFooter");
            }

            /// <summary>
            /// Append the cell contents we have collected.
            /// </summary>
            public void AppendCellText(StringBuilder buffer)
            {
                eb.CheckMaxTextSize(buffer, output.ToString());
                buffer.Append(output);
            }

            /// <summary>
            /// Reset this <c>SheetTextExtractor</c> for the next sheet.
            /// </summary>
            public void Reset()
            {
                output.Length = 0;
                firstCellOfRow = true;
                if(headerFooterMap != null)
                {
                    headerFooterMap.Clear();
                }
            }
        }
    }
}


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
using NPOI.OpenXml4Net.OPC;
using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using System.IO;
using System.Text;
namespace NPOI.XSSF.Extractor
{


    /**
     * Implementation of a text extractor from OOXML Excel
     *  files that uses SAX event based parsing.
     */
    public class XSSFEventBasedExcelExtractor : POIXMLTextExtractor
    {
        private OPCPackage container;
        private POIXMLProperties properties;

        private Locale locale;
        private bool includeSheetNames = true;
        private bool formulasNotResults = false;

        public XSSFEventBasedExcelExtractor(String path)
            : this(OPCPackage.Open(path))
        {

        }
        public XSSFEventBasedExcelExtractor(OPCPackage Container)
            : base(null)
        {

            this.container = Container;

            properties = new POIXMLProperties(Container);
        }

        /**
         * Should sheet names be included? Default is true
         */
        public void SetIncludeSheetNames(bool includeSheetNames)
        {
            this.includeSheetNames = includeSheetNames;
        }
        /**
         * Should we return the formula itself, and not
         *  the result it produces? Default is false
         */
        public void SetFormulasNotResults(bool formulasNotResults)
        {
            this.formulasNotResults = formulasNotResults;
        }

        public void SetLocale(Locale locale)
        {
            this.locale = locale;
        }

        /**
         * Returns the opened OPCPackage Container.
         */

        public OPCPackage GetPackage()
        {
            return container;
        }

        /**
         * Returns the core document properties
         */

        public NPOI.POIXMLProperties.CoreProperties GetCoreProperties()
        {
            return properties.GetCoreProperties();
        }
        /**
         * Returns the extended document properties
         */

        public NPOI.POIXMLProperties.ExtendedProperties GetExtendedProperties()
        {
            return properties.GetExtendedProperties();
        }
        /**
         * Returns the custom document properties
         */

        public NPOI.POIXMLProperties.CustomProperties GetCustomProperties()
        {
            return properties.GetCustomProperties();
        }

        /**
         * Processes the given sheet
         */
        public void ProcessSheet(
                SheetContentsHandler sheetContentsExtractor,
                StylesTable styles,
                ReadOnlySharedStringsTable strings,
                InputStream sheetInputStream)
        {

            DataFormatter formatter;
            if (locale == null)
            {
                formatter = new DataFormatter();
            }
            else
            {
                formatter = new DataFormatter(locale);
            }

            InputSource sheetSource = new InputSource(sheetInputStream);
            SAXParserFactory saxFactory = SAXParserFactory.newInstance();
            try
            {
                SAXParser saxParser = saxFactory.newSAXParser();
                XMLReader sheetParser = saxParser.GetXMLReader();
                ContentHandler handler = new XSSFSheetXMLHandler(
                      styles, strings, sheetContentsExtractor, formatter, formulasNotResults);
                sheetParser.SetContentHandler(handler);
                sheetParser.Parse(sheetSource);
            }
            catch (ParserConfigurationException e)
            {
                throw new RuntimeException("SAX Parser appears to be broken - " + e.GetMessage());
            }
        }

        /**
         * Processes the file and returns the text
         */
        public String GetText()
        {
            try
            {
                ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(container);
                XSSFReader xssfReader = new XSSFReader(container);
                StylesTable styles = xssfReader.GetStylesTable();
                XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator)xssfReader.GetSheetsData();

                StringBuilder text = new StringBuilder();
                SheetTextExtractor sheetExtractor = new SheetTextExtractor(text);

                while (iter.HasNext())
                {
                    InputStream stream = iter.next();
                    if (includeSheetNames)
                    {
                        text.Append(iter.GetSheetName());
                        text.Append('\n');
                    }
                    ProcessSheet(sheetExtractor, styles, strings, stream);
                    stream.Close();
                }

                return text.ToString();
            }
            catch (IOException e)
            {
                System.err.println(e);
                return null;
            }
            catch (OpenXML4NetException o4je)
            {
                System.err.println(o4je);
                return null;
            }
        }
        public void Close()
        {
            if (container != null)
            {
                container.Close();
                container = null;
            }
            base.close();
        }
        protected class SheetTextExtractor : SheetContentsHandler
        {
            private StringBuilder output;
            private bool firstCellOfRow = true;

            protected SheetTextExtractor(StringBuilder output)
            {
                this.output = output;
            }

            public void startRow(int rowNum)
            {
                firstCellOfRow = true;
            }

            public void endRow()
            {
                output.Append('\n');
            }

            public void cell(String cellRef, String formattedValue)
            {
                if (firstCellOfRow)
                {
                    firstCellOfRow = false;
                }
                else
                {
                    output.Append('\t');
                }
                output.Append(formattedValue);
            }

            public void headerFooter(String text, bool IsHeader, String tagName)
            {
                // We don't include headers in the output yet, so ignore
            }
        }
    }
}


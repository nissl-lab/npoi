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

using NPOI.OpenXml4Net;
using NPOI.SS.Extractor;
using NSAX;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XSSF.Extractor
{

    using NPOI;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF.Binary;
    using NPOI.XSSF.EventUserModel;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// Implementation of a text extractor or xlsb Excel
    /// files that uses SAX-like binary parsing.
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBEventBasedExcelExtractor : XSSFEventBasedExcelExtractor
            , IExcelExtractor
    {

        //private static  POILogger LOGGER = POILogFactory.GetLogger(XSSFBEventBasedExcelExtractor.class);

        public static  XSSFRelation[] SUPPORTED_TYPES = new XSSFRelation[]{
            XSSFRelation.XLSB_BINARY_WORKBOOK
    };

        private bool handleHyperlinksInCells = false;

        public XSSFBEventBasedExcelExtractor(String path)
            : base(path)
        {


        }

        public XSSFBEventBasedExcelExtractor(OPCPackage container)
        : base(container)
        {

            ;
        }

        public static void main(String[] args)
        {

            if(args.Length < 1)
            {
                //System.err.println("Use:");
                //System.err.println("  XSSFBEventBasedExcelExtractor <filename.xlsb>");
                //System.exit(1);
                return;
            }
            XSSFBEventBasedExcelExtractor extractor =
                new XSSFBEventBasedExcelExtractor(args[0]);
            Console.WriteLine(extractor.Text);
            extractor.Close();
        }

        public void SetHandleHyperlinksInCells(bool handleHyperlinksInCells)
        {
            this.handleHyperlinksInCells = handleHyperlinksInCells;
        }

        /// <summary>
        /// Should we return the formula itself, and not
        /// the result it produces? Default is false
        /// This is currently unsupported for xssfb
        /// </summary>
        public void SetFormulasNotResults(bool formulasNotResults)
        {
            throw new ArgumentException("Not currently supported");
        }

        /// <summary>
        /// Processes the given sheet
        /// </summary>
        public void ProcessSheet(
                XSSFSheetXMLHandler.ISheetContentsHandler sheetContentsExtractor,
                XSSFBStylesTable styles,
                XSSFBCommentsTable comments,
                XSSFBSharedStringsTable strings,
                Stream sheetInputStream)
        {
            DataFormatter formatter;
            if(Locale == null)
            {
                formatter = new DataFormatter();
            }
            else
            {
                formatter = new DataFormatter(Locale);
            }

            XSSFBSheetHandler xssfbSheetHandler = new XSSFBSheetHandler(
                sheetInputStream,
                styles, comments, strings, sheetContentsExtractor, formatter, FormulasNotResults
            );
            xssfbSheetHandler.Parse();
        }

        /// <summary>
        /// Processes the file and returns the text
        /// </summary>
        public String GetText()
        {
            try
            {
                XSSFBSharedStringsTable strings = new XSSFBSharedStringsTable(Package);
                XSSFBReader xssfbReader = new XSSFBReader(Package);
                XSSFBStylesTable styles = xssfbReader.GetXSSFBStylesTable();
                XSSFBReader.SheetIterator iter = (XSSFBReader.SheetIterator) xssfbReader.GetSheetsData();

                StringBuilder text = new StringBuilder();
                SheetTextExtractor sheetExtractor = new SheetTextExtractor(this);
                XSSFBHyperlinksTable hyperlinksTable = null;
                while(iter.MoveNext())
                {
                    Stream stream = iter.Current;
                    if(IncludeSheetNames)
                    {
                        text.Append(iter.SheetName);
                        text.Append('\n');
                    }
                    if(handleHyperlinksInCells)
                    {
                        hyperlinksTable = new XSSFBHyperlinksTable(iter.SheetPart);
                    }
                    XSSFBCommentsTable comments = IncludeCellComments ? iter.GetXSSFBSheetComments() : null;
                    ProcessSheet(sheetExtractor, styles, comments, strings, stream);
                    if(IncludeHeadersFooters)
                    {
                        sheetExtractor.AppendHeaderText(text);
                    }
                    sheetExtractor.AppendCellText(text);
                    if(IncludeTextBoxes)
                    {
                        ProcessShapes(iter.Shapes, text);
                    }
                    if(IncludeHeadersFooters)
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
}


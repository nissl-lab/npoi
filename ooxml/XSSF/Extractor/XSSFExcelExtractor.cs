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
using NPOI.XSSF.UserModel;
using System;
using NPOI.OpenXml4Net.OPC;
using System.Text;
using NPOI.SS.UserModel;
using System.Collections;
using System.Globalization;
namespace NPOI.XSSF.Extractor
{
    /**
     * Helper class to extract text from an OOXML Excel file
     */
    public class XSSFExcelExtractor : POIXMLTextExtractor, NPOI.SS.Extractor.IExcelExtractor
    {
        public static XSSFRelation[] SUPPORTED_TYPES = new XSSFRelation[] {
      XSSFRelation.WORKBOOK, XSSFRelation.MACRO_TEMPLATE_WORKBOOK,
      XSSFRelation.MACRO_ADDIN_WORKBOOK, XSSFRelation.TEMPLATE_WORKBOOK,
      XSSFRelation.MACROS_WORKBOOK
   };

        private XSSFWorkbook workbook;
        private bool includeSheetNames = true;
        private bool formulasNotResults = false;
        private bool includeCellComments = false;
        private bool includeHeadersFooters = true;

        public XSSFExcelExtractor(String path)
            : this(new XSSFWorkbook(path))
        {

        }
        public XSSFExcelExtractor(OPCPackage Container)
            : this(new XSSFWorkbook(Container))
        {

        }
        public XSSFExcelExtractor(XSSFWorkbook workbook)
            : base(workbook)
        {

            this.workbook = workbook;
        }
        /// <summary>
        ///  Should header and footer be included? Default is true
        /// </summary>
        public bool IncludeHeaderFooter
        {
            get
            {
                return this.includeHeadersFooters;
            }
            set
            {
                this.includeHeadersFooters = value;
            }
        }
        /// <summary>
        /// Should sheet names be included? Default is true
        /// </summary>
        /// <value>if set to <c>true</c> [include sheet names].</value>
        public bool IncludeSheetNames
        {
            get
            {
                return this.includeSheetNames;
            }
            set
            {
                this.includeSheetNames = value;
            }
        }
        /// <summary>
        /// Should we return the formula itself, and not
        /// the result it produces? Default is false
        /// </summary>
        /// <value>if set to <c>true</c> [formulas not results].</value>
        public bool FormulasNotResults
        {
            get
            {
                return this.formulasNotResults;
            }
            set
            {
                this.formulasNotResults = value;
            }
        }
        /// <summary>
        /// Should cell comments be included? Default is false
        /// </summary>
        /// <value>if set to <c>true</c> [include cell comments].</value>
        public bool IncludeCellComments
        {
            get
            {
                return this.includeCellComments;
            }
            set
            {
                this.includeCellComments = value;
            }
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
        /**
         * Should cell comments be included? Default is false
         */
        public void SetIncludeCellComments(bool includeCellComments)
        {
            this.includeCellComments = includeCellComments;
        }
        /**
         * Should headers and footers be included? Default is true
         */
        public void SetIncludeHeadersFooters(bool includeHeadersFooters)
        {
            this.includeHeadersFooters = includeHeadersFooters;
        }
        public void SetLocale(CultureInfo locale) {
            this.locale = locale;
        }

        private CultureInfo locale=null;
        /**
         * Retreives the text contents of the file
         */
        public override String Text
        {
            get
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

                StringBuilder text = new StringBuilder();

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(i);
                    if (includeSheetNames)
                    {
                        text.Append(workbook.GetSheetName(i)+"\n");
                    }

                    // Header(s), if present
                    if (includeHeadersFooters)
                    {
                        text.Append(
                                ExtractHeaderFooter(sheet.FirstHeader)
                        );
                        text.Append(
                                ExtractHeaderFooter(sheet.OddHeader)
                        );
                        text.Append(
                                ExtractHeaderFooter(sheet.EvenHeader)
                        );
                    }

                    // Rows and cells
                    foreach (Object rawR in sheet)
                    {
                        IRow row = (IRow)rawR;
                        IEnumerator ri =row.GetEnumerator();
                        bool firsttime=true;
                        while (ri.MoveNext())
                        {
                            if (!firsttime)
                            {
                                text.Append("\t");
                            }
                            else
                            {
                                firsttime = false;
                            }
                            ICell cell = (ICell)ri.Current;

                            // Is it a formula one?
                            if (cell.CellType == CellType.Formula)
                            {
                                if (formulasNotResults)
                                {
                                    text.Append(cell.CellFormula);
                                }
                                else
                                {
                                    if (cell.CachedFormulaResultType == CellType.String)
                                    {
                                        HandleStringCell(text, cell);
                                    }
                                    else
                                    {
                                        HandleNonStringCell(text, cell, formatter);
                                    }
                                }

                            }
                            else if (cell.CellType == CellType.String)
                            {
                                HandleStringCell(text, cell);
                            }
                            else
                            {
                                HandleNonStringCell(text, cell, formatter);
                            }

                            // Output the comment, if requested and exists
                            IComment comment = cell.CellComment;
                            if (includeCellComments && comment != null)
                            {
                                // Replace any newlines with spaces, otherwise it
                                //  breaks the output
                                String commentText = comment.String.String.Replace('\n', ' ');
                                text.Append(" Comment by ").Append(comment.Author).Append(": ").Append(commentText);
                            }
                            
                        }
                        text.Append("\n");
                    }

                    // Finally footer(s), if present
                    if (includeHeadersFooters)
                    {
                        text.Append(
                                ExtractHeaderFooter(sheet.FirstFooter)
                        );
                        text.Append(
                                ExtractHeaderFooter(sheet.OddFooter)
                        );
                        text.Append(
                                ExtractHeaderFooter(sheet.EvenFooter)
                        );
                    }
                }

                return text.ToString();
            }
        }
        private void HandleStringCell(StringBuilder text, ICell cell)
        {
            text.Append(cell.RichStringCellValue.String);
        }
        private void HandleNonStringCell(StringBuilder text, ICell cell, DataFormatter formatter)
        {
            CellType type = cell.CellType;
            if (type == CellType.Formula)
            {
                type = cell.CachedFormulaResultType;
            }

            if (type == CellType.Numeric)
            {
                ICellStyle cs = cell.CellStyle;

                if (cs.GetDataFormatString()!= null)
                {
                    text.Append(formatter.FormatRawCellContents(
                          cell.NumericCellValue, cs.DataFormat, cs.GetDataFormatString()
                    ));
                    return;
                }
            }

            // No supported styling applies to this cell
            XSSFCell xcell = (XSSFCell)cell;
            text.Append(xcell.GetRawValue());
        }
        private String ExtractHeaderFooter(IHeaderFooter hf)
        {
            return NPOI.HSSF.Extractor.ExcelExtractor.ExtractHeaderFooter(hf);
        }
    }

}
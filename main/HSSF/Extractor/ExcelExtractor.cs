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

namespace NPOI.HSSF.Extractor
{
    using System;
    using System.Text;

    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NPOI.SS.Extractor;

    /// <summary>
    /// A text extractor for Excel files.
    /// Returns the textual content of the file, suitable for
    /// indexing by something like Lucene, but not really
    /// intended for display to the user.
    /// </summary>
    public class ExcelExtractor : POIOLE2TextExtractor, IExcelExtractor
    {
        private HSSFWorkbook wb;
        private HSSFDataFormatter _formatter;
        private bool includeSheetNames = true;
        private bool formulasNotResults = false;
        private bool includeCellComments = false;
        private bool includeBlankCells = false;
        private bool includeHeaderFooter = true;
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExtractor"/> class.
        /// </summary>
        /// <param name="wb">The wb.</param>
        public ExcelExtractor(HSSFWorkbook wb)
            : base(wb)
        {
            this.wb = wb;
            _formatter = new HSSFDataFormatter();
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExtractor"/> class.
        /// </summary>
        /// <param name="fs">The fs.</param>
        public ExcelExtractor(POIFSFileSystem fs)
            : this(new HSSFWorkbook(fs))
        {
        }

        /// <summary>
        ///  Should header and footer be included? Default is true
        /// </summary>
        public bool IncludeHeaderFooter
        {
            get {
                return this.includeHeaderFooter;
            }
            set {
                this.includeHeaderFooter = value;
            }
        }
        /// <summary>
        /// Should sheet names be included? Default is true
        /// </summary>
        /// <value>if set to <c>true</c> [include sheet names].</value>
        public bool IncludeSheetNames
        {
            get {
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
        /// <summary>
        /// Should blank cells be output? Default is to only
        /// output cells that are present in the file and are
        /// non-blank.
        /// </summary>
        /// <value>if set to <c>true</c> [include blank cells].</value>
        public bool IncludeBlankCells
        {
            get
            {
                return this.includeBlankCells;
            }
            set
            {
                this.includeBlankCells = value;
            }
        }

        /// <summary>
        /// Retreives the text contents of the file
        /// </summary>
        /// <value>All the text from the document.</value>
        public override String Text
        {
            get
            {
                StringBuilder text = new StringBuilder();

                // We don't care about the differnce between
                //  null (missing) and blank cells
                wb.MissingCellPolicy = MissingCellPolicy.RETURN_BLANK_AS_NULL;

                // Process each sheet in turn
                for (int i = 0; i < wb.NumberOfSheets; i++)
                {
                    HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(i);
                    if (sheet == null) { continue; }

                    if (includeSheetNames)
                    {
                        String name = wb.GetSheetName(i);
                        if (name != null)
                        {
                            text.Append(name);
                            text.Append("\n");
                        }
                    }

                    // Header text, if there is any
                    if (sheet.Header != null && includeHeaderFooter)
                    {
                        text.Append(
                                ExtractHeaderFooter(sheet.Header)
                        );
                    }

                    int firstRow = sheet.FirstRowNum;
                    int lastRow = sheet.LastRowNum;
                    for (int j = firstRow; j <= lastRow; j++)
                    {
                        IRow row = sheet.GetRow(j);
                        if (row == null) { continue; }

                        // Check each cell in turn
                        int firstCell = row.FirstCellNum;
                        int lastCell = row.LastCellNum;
                        if (includeBlankCells)
                        {
                            firstCell = 0;
                        }

                        for (int k = firstCell; k < lastCell; k++)
                        {
                            ICell cell = row.GetCell(k);
                            bool outputContents = true;

                            if (cell == null)
                            {
                                // Only output if requested
                                outputContents = includeBlankCells;
                            }
                            else
                            {
                                switch (cell.CellType)
                                {
                                    case CellType.String:
                                        text.Append(cell.RichStringCellValue.String);
                                        break;
                                    case CellType.Numeric:
                                        // Note - we don't apply any formatting!
                                        //text.Append(cell.NumericCellValue);
                                        text.Append(_formatter.FormatCellValue(cell));
                                        break;
                                    case CellType.Boolean:
                                        text.Append(cell.BooleanCellValue);
                                        break;
                                    case CellType.Error:
                                        text.Append(ErrorEval.GetText(cell.ErrorCellValue));
                                        break;
                                    case CellType.Formula:
                                        if (formulasNotResults)
                                        {
                                            text.Append(cell.CellFormula);
                                        }
                                        else
                                        {
                                            switch (cell.CachedFormulaResultType)
                                            {
                                                case CellType.String:
                                                    IRichTextString str = cell.RichStringCellValue;
                                                    if (str != null && str.Length > 0)
                                                    {
                                                        text.Append(str.ToString());
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    //text.Append(cell.NumericCellValue);
                                                    HSSFCellStyle style = (HSSFCellStyle)cell.CellStyle;
                                                    if (style == null)
                                                    {
                                                        text.Append(cell.NumericCellValue);
                                                    }
                                                    else
                                                    {
                                                        text.Append(
                                                              _formatter.FormatRawCellContents(
                                                                    cell.NumericCellValue,
                                                                    style.DataFormat,
                                                                    style.GetDataFormatString()
                                                              )
                                                        );
                                                    }
                                                    break;
                                                case CellType.Boolean:
                                                    text.Append(cell.BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    text.Append(ErrorEval.GetText(cell.ErrorCellValue));
                                                    break;

                                            }
                                        }
                                        break;
                                    default:
                                        throw new Exception("Unexpected cell type (" + cell.CellType + ")");
                                }

                                // Output the comment, if requested and exists
                                NPOI.SS.UserModel.IComment comment = cell.CellComment;
                                if (includeCellComments && comment != null)
                                {
                                    // Replace any newlines with spaces, otherwise it
                                    //  breaks the output
                                    String commentText = comment.String.String.Replace('\n', ' ');
                                    text.Append(" Comment by " + comment.Author + ": " + commentText);
                                }
                            }

                            // Output a tab if we're not on the last cell
                            if (outputContents && k < (lastCell - 1))
                            {
                                text.Append("\t");
                            }
                        }

                        // Finish off the row
                        text.Append("\n");
                    }

                    // Finally Feader text, if there is any
                    if (sheet.Footer != null && includeHeaderFooter)
                    {
                        text.Append(
                                ExtractHeaderFooter(sheet.Footer)
                        );
                    }
                }

                return text.ToString();
            }
        }

        /// <summary>
        /// Extracts the header footer.
        /// </summary>
        /// <param name="hf">The header or footer</param>
        /// <returns></returns>
        public static String ExtractHeaderFooter(NPOI.SS.UserModel.IHeaderFooter hf)
        {
            StringBuilder text = new StringBuilder();

            if (hf.Left != null)
            {
                text.Append(hf.Left);
            }
            if (hf.Center != null)
            {
                if (text.Length > 0)
                    text.Append("\t");
                text.Append(hf.Center);
            }
            if (hf.Right != null)
            {
                if (text.Length > 0)
                    text.Append("\t");
                text.Append(hf.Right);
            }
            if (text.Length > 0)
                text.Append("\n");

            return text.ToString();
        }
    }
}
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
    using System.IO;
    using System.Collections;

    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;
    using NPOI;
    using NPOI.HPSF;
    using NPOI.HSSF.EventUserModel;
    using NPOI.HSSF.Model;
    //using NPOI.HSSF.Util;
    using NPOI.SS.Util;
    using System.Globalization;

    /// <summary>
    /// A text extractor for Excel files, that is based
    /// on the hssf eventusermodel api.
    /// It will typically use less memory than
    /// ExcelExtractor, but may not provide
    /// the same richness of formatting.
    /// Returns the textual content of the file, suitable for
    /// indexing by something like Lucene, but not really
    /// intended for display to the user.
    /// </summary>
    public class EventBasedExcelExtractor : POIOLE2TextExtractor
    {
        private POIFSFileSystem fs;
        private bool includeSheetNames = true;
        private bool formulasNotResults = false;

        public EventBasedExcelExtractor(POIFSFileSystem fs)
            : base(null)
        {

            this.fs = fs;
        }

        /// <summary>
        /// Would return the document information metadata for the document,
        /// if we supported it
        /// </summary>
        /// <value>The doc summary information.</value>
        public override DocumentSummaryInformation DocSummaryInformation
        {
            get { 
                throw new NotImplementedException("Metadata extraction not supported in streaming mode, please use ExcelExtractor"); 
            }
        }
        /// <summary>
        /// Would return the summary information metadata for the document,
        /// if we supported it
        /// </summary>
        /// <value>The summary information.</value>
        public override SummaryInformation SummaryInformation
        {
            get
            {
                throw new NotImplementedException("Metadata extraction not supported in streaming mode, please use ExcelExtractor");
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
        /// Retreives the text contents of the file
        /// </summary>
        /// <value>All the text from the document.</value>
        public override String Text
        {
            get
            {
                String text = null;
                try
                {
                    TextListener tl = TriggerExtraction();

                    text = tl.text.ToString();
                    if (!text.EndsWith("\n", StringComparison.Ordinal))
                    {
                        text = text + "\n";
                    }
                }
                catch (IOException)
                {
                    throw;
                }

                return text;
            }
        }

        /// <summary>
        /// Triggers the extraction.
        /// </summary>
        /// <returns></returns>
        private TextListener TriggerExtraction()
        {
            TextListener tl = new TextListener(includeSheetNames,formulasNotResults);
            FormatTrackingHSSFListener ft = new FormatTrackingHSSFListener(tl);
            tl.ft = ft;

            // Register and process
            HSSFEventFactory factory = new HSSFEventFactory();
            HSSFRequest request = new HSSFRequest();
            request.AddListenerForAllRecords(ft);

            factory.ProcessWorkbookEvents(request, fs);

            return tl;
        }

        private class TextListener : IHSSFListener
        {
            public FormatTrackingHSSFListener ft;
            private SSTRecord sstRecord;

            private IList sheetNames = new ArrayList();
            public StringBuilder text = new StringBuilder();
            private int sheetNum = -1;
            private int rowNum;

            private bool outputNextStringValue = false;
            private int nextRow = -1;

            private bool includeSheetNames;
            private bool formulasNotResults;

            public TextListener(bool includeSheetNames, bool formulasNotResults)
            {
                this.includeSheetNames = includeSheetNames;
                this.formulasNotResults = formulasNotResults;
            }

            /// <summary>
            /// Process an HSSF Record. Called when a record occurs in an HSSF file.
            /// </summary>
            /// <param name="record"></param>
            public void ProcessRecord(Record record)
            {
                String thisText = null;
                int thisRow = -1;

                switch (record.Sid)
                {
                    case BoundSheetRecord.sid:
                        BoundSheetRecord sr = (BoundSheetRecord)record;
                        sheetNames.Add(sr.Sheetname);
                        break;
                    case BOFRecord.sid:
                        BOFRecord bof = (BOFRecord)record;
                        if (bof.Type == BOFRecordType.Worksheet)
                        {
                            sheetNum++;
                            rowNum = -1;

                            if (includeSheetNames)
                            {
                                if (text.Length > 0) text.Append("\n");
                                text.Append(sheetNames[sheetNum]);
                            }
                        }
                        break;
                    case SSTRecord.sid:
                        sstRecord = (SSTRecord)record;
                        break;

                    case FormulaRecord.sid:
                        FormulaRecord frec = (FormulaRecord)record;
                        thisRow = frec.Row;

                        if (formulasNotResults)
                        {
                            thisText = HSSFFormulaParser.ToFormulaString((HSSFWorkbook)null, frec.ParsedExpression);
                        }
                        else
                        {
                            if (frec.HasCachedResultString)
                            {
                                // Formula result is a string
                                // This is stored in the next record
                                outputNextStringValue = true;
                                nextRow = frec.Row;
                            }
                            else
                            {
                                thisText = FormatNumberDateCell(frec, frec.Value);
                            }
                        }
                        break;
                    case StringRecord.sid:
                        if (outputNextStringValue)
                        {
                            // String for formula
                            StringRecord srec = (StringRecord)record;
                            thisText = srec.String;
                            thisRow = nextRow;
                            outputNextStringValue = false;
                        }
                        break;
                    case LabelRecord.sid:
                        LabelRecord lrec = (LabelRecord)record;
                        thisRow = lrec.Row;
                        thisText = lrec.Value;
                        break;
                    case LabelSSTRecord.sid:
                        LabelSSTRecord lsrec = (LabelSSTRecord)record;
                        thisRow = lsrec.Row;
                        if (sstRecord == null)
                        {
                            throw new Exception("No SST record found");
                        }
                        thisText = sstRecord.GetString(lsrec.SSTIndex).ToString();
                        break;
                    case NoteRecord.sid:
                        NoteRecord nrec = (NoteRecord)record;
                        thisRow = nrec.Row;
                        // TODO: Find object to match nrec.GetShapeId()
                        break;
                    case NumberRecord.sid:
                        NumberRecord numrec = (NumberRecord)record;
                        thisRow = numrec.Row;
                        thisText = FormatNumberDateCell(numrec, numrec.Value);
                        break;
                    default:
                        break;
                }

                if (thisText != null)
                {
                    if (thisRow != rowNum)
                    {
                        rowNum = thisRow;
                        if (text.Length > 0)
                            text.Append("\n");
                    }
                    else
                    {
                        text.Append("\t");
                    }
                    text.Append(thisText);
                }
            }

            /// <summary>
            /// Formats a number or date cell, be that a real number, or the
            /// answer to a formula
            /// </summary>
            /// <param name="cell">The cell.</param>
            /// <param name="value">The value.</param>
            /// <returns></returns>
            private String FormatNumberDateCell(CellValueRecordInterface cell, double value)
            {
                // Get the built in format, if there is one
                int formatIndex = ft.GetFormatIndex(cell);
                String formatString = ft.GetFormatString(cell);

                if (formatString == null)
                {
                    return value.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    // Is it a date?
                    if (NPOI.SS.UserModel.DateUtil.IsADateFormat(formatIndex, formatString) &&
                            NPOI.SS.UserModel.DateUtil.IsValidExcelDate(value))
                    {
                        // Java wants M not m for month
                        formatString = formatString.Replace('m', 'M');
                        // Change \- into -, if it's there
                        formatString = formatString.Replace("\\\\-", "-");

                        // Format as a date
                        DateTime d = NPOI.SS.UserModel.DateUtil.GetJavaDate(value, false);
                        SimpleDateFormat df = new SimpleDateFormat(formatString);
                        return df.Format(d, CultureInfo.CurrentCulture);
                    }
                    else
                    {
                        if (formatString == "General")
                        {
                            // Some sort of wierd default
                            return value.ToString(CultureInfo.InvariantCulture);
                        }

                        // Format as a number
                        DecimalFormat df = new DecimalFormat(formatString);
                        return df.Format(value, CultureInfo.CurrentCulture);
                    }
                }
            }
        }
    }
}
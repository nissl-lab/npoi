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

namespace NPOI.XSSF.EventUserModel
{
    using static NPOI.XSSF.UserModel.XSSFRelation;


    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.UserModel;
    using NSAX.Helpers;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NSAX;

    /// <summary>
    /// This class handles the processing of a sheet#.xml
    ///  sheet part of a XSSF .xlsx file, and generates
    ///  row and cell events for it.
    /// </summary>
    public class XSSFSheetXMLHandler : DefaultHandler
    {
        //private static  POILogger logger = POILogFactory.GetLogger(XSSFSheetXMLHandler.class);

        /// <summary>
        /// These are the different kinds of cells we support.
        /// We keep track of the current one between
        ///  the start and end.
        /// </summary>
        enum XSSFDataType
        {
            Boolean,
            Error,
            Formula,
            InlineString,
            SSTString,
            Number,
        }

        /// <summary>
        /// Table with the styles used for formatting
        /// </summary>
        private StylesTable stylesTable;

        /// <summary>
        /// Table with cell comments
        /// </summary>
        private CommentsTable commentsTable;

        /// <summary>
        /// Read only access to the shared strings table, for looking
        ///  up (most) string cell's contents
        /// </summary>
        private ReadOnlySharedStringsTable sharedStringsTable;

        /// <summary>
        /// Where our text is going
        /// </summary>
        private ISheetContentsHandler output;

        // Set when V start element is seen
        private bool vIsOpen;
        // Set when F start element is seen
        private bool fIsOpen;
        // Set when an Inline String "is" is seen
        private bool isIsOpen;
        // Set when a header/footer element is seen
        private bool hfIsOpen;

        // Set when cell start element is seen;
        // used when cell close element is seen.
        private XSSFDataType nextDataType;

        // Used to format numeric cell values.
        private short formatIndex;
        private String formatString;
        private  DataFormatter formatter;
        private int rowNum;
        private int nextRowNum;      // some sheets do not have rowNums, Excel can read them so we should try to handle them correctly as well
        private String cellRef;
        private bool formulasNotResults;

        // Gathers characters as they are seen.
        private StringBuilder value = new StringBuilder();
        private StringBuilder formula = new StringBuilder();
        private StringBuilder headerFooter = new StringBuilder();

        private Queue<CellAddress> commentCellRefs;

        /// <summary>
        /// Accepts objects needed while parsing.
        /// </summary>
        /// <param name="styles"> Table of styles</param>
        /// <param name="strings">Table of shared strings</param>
        public XSSFSheetXMLHandler(
                StylesTable styles,
                CommentsTable comments,
                ReadOnlySharedStringsTable strings,
                ISheetContentsHandler sheetContentsHandler,
                DataFormatter dataFormatter,
                bool formulasNotResults)
        {
            this.stylesTable = styles;
            this.commentsTable = comments;
            this.sharedStringsTable = strings;
            this.output = sheetContentsHandler;
            this.formulasNotResults = formulasNotResults;
            this.nextDataType = XSSFDataType.Number;
            this.formatter = dataFormatter;
            Init();
        }

        /// <summary>
        /// Accepts objects needed while parsing.
        /// </summary>
        /// <param name="styles"> Table of styles</param>
        /// <param name="strings">Table of shared strings</param>
        public XSSFSheetXMLHandler(
                StylesTable styles,
                ReadOnlySharedStringsTable strings,
                ISheetContentsHandler sheetContentsHandler,
                DataFormatter dataFormatter,
                bool formulasNotResults)
                 : this(styles, null, strings, sheetContentsHandler, dataFormatter, formulasNotResults)
        {

        }

        /// <summary>
        /// Accepts objects needed while parsing.
        /// </summary>
        /// <param name="styles"> Table of styles</param>
        /// <param name="strings">Table of shared strings</param>
        public XSSFSheetXMLHandler(
                StylesTable styles,
                ReadOnlySharedStringsTable strings,
                ISheetContentsHandler sheetContentsHandler,
                bool formulasNotResults)
                 : this(styles, strings, sheetContentsHandler, new DataFormatter(), formulasNotResults)
        {

        }

        private void Init()
        {
            if(commentsTable != null)
            {
                commentCellRefs = new Queue<CellAddress>();
                //noinspection deprecation
                foreach(CT_Comment comment in commentsTable.GetCTComments().commentList.GetCommentArray())
                {
                    commentCellRefs.Enqueue(new CellAddress(comment.@ref));
                }
            }
        }

        private bool IsTextTag(String name)
        {
            if("v".Equals(name))
            {
                // Easy, normal v text tag
                return true;
            }
            if("inlineStr".Equals(name))
            {
                // Easy inline string
                return true;
            }
            if("t".Equals(name) && isIsOpen)
            {
                // Inline string <is><t>...</t></is> pair
                return true;
            }
            // It isn't a text tag
            return false;
        }
        public override void StartElement(String uri, String localName, String qName,
                                 IAttributes attributes)
        {


            if(uri != null && !uri.Equals(NS_SPREADSHEETML))
            {
                return;
            }

            if(IsTextTag(localName))
            {
                vIsOpen = true;
                // Clear contents cache
                value.Length = 0;
            }
            else if("is".Equals(localName))
            {
                // Inline string outer tag
                isIsOpen = true;
            }
            else if("f".Equals(localName))
            {
                // Clear contents cache
                formula.Length = 0;

                // Mark us as being a formula if not already
                if(nextDataType == XSSFDataType.Number)
                {
                    nextDataType = XSSFDataType.Formula;
                }

                // Decide where to Get the formula string from
                String type = attributes.GetValue("t");
                if(type != null && type.Equals("shared"))
                {
                    // Is it the one that defines the shared, or uses it?
                    String ref1 = attributes.GetValue("ref");
                    String si = attributes.GetValue("si");

                    if(ref1 != null)
                    {
                        // This one defines it
                        // TODO Save it somewhere
                        fIsOpen = true;
                    }
                    else
                    {
                        // This one uses a shared formula
                        // TODO Retrieve the shared formula and tweak it to 
                        //  match the current cell
                        if(formulasNotResults)
                        {
                            //logger.log(POILogger.WARN, "shared formulas not yet supported!");
                        } 
                        /*else {
                           // It's a shared formula, so we can't Get at the formula string yet
                           // However, they don't care about the formula string, so that's ok!
                        }*/
                    }
                }
                else
                {
                    fIsOpen = true;
                }
            }
            else if("oddHeader".Equals(localName) || "evenHeader".Equals(localName) ||
                  "firstHeader".Equals(localName) || "firstFooter".Equals(localName) ||
                  "oddFooter".Equals(localName) || "evenFooter".Equals(localName))
            {
                hfIsOpen = true;
                // Clear contents cache
                headerFooter.Length = 0;
            }
            else if("row".Equals(localName))
            {
                String rowNumStr = attributes.GetValue("r");
                if(rowNumStr != null)
                {
                    rowNum = Int32.Parse(rowNumStr) - 1;
                }
                else
                {
                    rowNum = nextRowNum;
                }
                output.StartRow(rowNum);
            }
            // c => cell
            else if("c".Equals(localName))
            {
                // Set up defaults.
                this.nextDataType = XSSFDataType.Number;
                this.formatIndex = -1;
                this.formatString = null;
                cellRef = attributes.GetValue("r");
                String cellType = attributes.GetValue("t");
                String cellStyleStr = attributes.GetValue("s");
                if("b".Equals(cellType))
                    nextDataType = XSSFDataType.Boolean;
                else if("e".Equals(cellType))
                    nextDataType = XSSFDataType.Error;
                else if("inlineStr".Equals(cellType))
                    nextDataType = XSSFDataType.InlineString;
                else if("s".Equals(cellType))
                    nextDataType = XSSFDataType.SSTString;
                else if("str".Equals(cellType))
                    nextDataType = XSSFDataType.Formula;
                else
                {
                    // Number, but almost certainly with a special style or format
                    XSSFCellStyle style = null;
                    if(stylesTable != null)
                    {
                        if(cellStyleStr != null)
                        {
                            int styleIndex = int.Parse(cellStyleStr);
                            style = stylesTable.GetStyleAt(styleIndex);
                        }
                        else if(stylesTable.NumCellStyles > 0)
                        {
                            style = stylesTable.GetStyleAt(0);
                        }
                    }
                    if(style != null)
                    {
                        this.formatIndex = style.DataFormat;
                        this.formatString = style.GetDataFormatString();
                        if(this.formatString == null)
                            this.formatString = BuiltinFormats.GetBuiltinFormat(this.formatIndex);
                    }
                }
            }
        }
        public override void EndElement(String uri, String localName, String qName)

        {


            if(uri != null && !uri.Equals(NS_SPREADSHEETML))
            {
                return;
            }

            String thisStr = null;

            // v => contents of a cell
            if(IsTextTag(localName))
            {
                vIsOpen = false;

                // Process the value contents as required, now we have it all
                switch(nextDataType)
                {
                    case XSSFDataType.Boolean:
                        char first = value[0];
                        thisStr = first == '0' ? "FALSE" : "TRUE";
                        break;

                    case XSSFDataType.Error:
                        thisStr = "ERROR:" + value;
                        break;

                    case XSSFDataType.Formula:
                        if(formulasNotResults)
                        {
                            thisStr = formula.ToString();
                        }
                        else
                        {
                            String fv = value.ToString();

                            if(this.formatString != null)
                            {
                                try
                                {
                                    // Try to use the value as a formattable number
                                    double d = double.Parse(fv);
                                    thisStr = formatter.FormatRawCellContents(d, this.formatIndex, this.formatString);
                                }
                                catch(FormatException)
                                {
                                    // Formula is a String result not a Numeric one
                                    thisStr = fv;
                                }
                            }
                            else
                            {
                                // No formatting applied, just do raw value in all cases
                                thisStr = fv;
                            }
                        }
                        break;

                    case XSSFDataType.InlineString:
                        // TODO: Can these ever have formatting on them?
                        XSSFRichTextString rtsi = new XSSFRichTextString(value.ToString());
                        thisStr = rtsi.ToString();
                        break;

                    case XSSFDataType.SSTString:
                        String sstIndex = value.ToString();
                        try
                        {
                            int idx = int.Parse(sstIndex);
                            XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.GetEntryAt(idx));
                            thisStr = rtss.ToString();
                        }
                        catch(FormatException)
                        {
                            //logger.log(POILogger.ERROR, "Failed to parse SST index '" + sstIndex, ex);
                        }
                        break;

                    case XSSFDataType.Number:
                        String n = value.ToString();
                        if(this.formatString != null && n.Length > 0)
                            thisStr = formatter.FormatRawCellContents(Double.Parse(n), this.formatIndex, this.formatString);
                        else
                            thisStr = n;
                        break;

                    default:
                        thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                        break;
                }

                // Do we have a comment for this cell?
                CheckForEmptyCellComments(EmptyCellCommentsCheckType.Cell);
                XSSFComment comment = commentsTable != null ? commentsTable.FindCellComment(new CellAddress(cellRef)) : null;

                // Output
                output.Cell(cellRef, thisStr, comment);
            }
            else if("f".Equals(localName))
            {
                fIsOpen = false;
            }
            else if("is".Equals(localName))
            {
                isIsOpen = false;
            }
            else if("row".Equals(localName))
            {
                // Handle any "missing" cells which had comments attached
                CheckForEmptyCellComments(EmptyCellCommentsCheckType.EndOfRow);

                // Finish up the row
                output.EndRow(rowNum);

                // some sheets do not have rowNum Set in the XML, Excel can read them so we should try to read them as well
                nextRowNum = rowNum + 1;
            }
            else if("sheetData".Equals(localName))
            {
                // Handle any "missing" cells which had comments attached
                CheckForEmptyCellComments(EmptyCellCommentsCheckType.EndOfSheetData);
            }
            else if("oddHeader".Equals(localName) || "evenHeader".Equals(localName) ||
                  "firstHeader".Equals(localName))
            {
                hfIsOpen = false;
                output.HeaderFooter(headerFooter.ToString(), true, localName);
            }
            else if("oddFooter".Equals(localName) || "evenFooter".Equals(localName) ||
                  "firstFooter".Equals(localName))
            {
                hfIsOpen = false;
                output.HeaderFooter(headerFooter.ToString(), false, localName);
            }
        }

        /// <summary>
        /// Captures characters only if a suitable element is open.
        /// Originally was just "v"; extended for inlineStr also.
        /// </summary>
        public override void Characters(char[] ch, int start, int length)

        {

            if(vIsOpen)
            {
                value.Append(ch, start, length);
            }
            if(fIsOpen)
            {
                formula.Append(ch, start, length);
            }
            if(hfIsOpen)
            {
                headerFooter.Append(ch, start, length);
            }
        }

        /// <summary>
        /// Do a check for, and output, comments in otherwise empty cells.
        /// </summary>
        private void CheckForEmptyCellComments(EmptyCellCommentsCheckType type)
        {
            if(commentCellRefs != null && commentCellRefs.Count>0)
            {
                // If we've reached the end of the sheet data, output any
                //  comments we haven't yet already handled
                if(type == EmptyCellCommentsCheckType.EndOfSheetData)
                {
                    while(commentCellRefs.Count>0)
                    {
                        OutputEmptyCellComment(commentCellRefs.Dequeue());
                    }
                    return;
                }

                // At the end of a row, handle any comments for "missing" rows before us
                if(this.cellRef == null)
                {
                    if(type == EmptyCellCommentsCheckType.EndOfRow)
                    {
                        while(commentCellRefs.Count>0)
                        {
                            if(commentCellRefs.Peek().Row == rowNum)
                            {
                                OutputEmptyCellComment(commentCellRefs.Dequeue());
                            }
                            else
                            {
                                return;
                            }
                        }
                        return;
                    }
                    else
                    {
                        throw new InvalidOperationException("Cell ref should be null only if there are only empty cells in the row; rowNum: " + rowNum);
                    }
                }

                CellAddress nextCommentCellRef;
                do
                {
                    CellAddress cellRef = new CellAddress(this.cellRef);
                    CellAddress peekCellRef = commentCellRefs.Peek();
                    if(type == EmptyCellCommentsCheckType.Cell && cellRef.Equals(peekCellRef))
                    {
                        // remove the comment cell ref from the list if we're about to handle it alongside the cell content
                        commentCellRefs.Dequeue();
                        return;
                    }
                    else
                    {
                        // fill in any gaps if there are empty cells with comment mixed in with non-empty cells
                        int comparison = peekCellRef.CompareTo(cellRef);
                        if(comparison > 0 && type == EmptyCellCommentsCheckType.EndOfRow && peekCellRef.Row <= rowNum)
                        {
                            nextCommentCellRef = commentCellRefs.Dequeue();
                            OutputEmptyCellComment(nextCommentCellRef);
                        }
                        else if(comparison < 0 && type == EmptyCellCommentsCheckType.Cell && peekCellRef.Row <= rowNum)
                        {
                            nextCommentCellRef = commentCellRefs.Dequeue();
                            OutputEmptyCellComment(nextCommentCellRef);
                        }
                        else
                        {
                            nextCommentCellRef = null;
                        }
                    }
                } while(nextCommentCellRef != null && commentCellRefs.Count>0);
            }
        }


        /// <summary>
        /// Output an empty-cell comment.
        /// </summary>
        private void OutputEmptyCellComment(CellAddress cellRef)
        {
            XSSFComment comment = commentsTable.FindCellComment(cellRef);
            output.Cell(cellRef.FormatAsString(), null, comment);
        }

        private enum EmptyCellCommentsCheckType
        {
            Cell,
            EndOfRow,
            EndOfSheetData
        }

        /// <summary>
        /// You need to implement this to handle the results
        ///  of the sheet parsing.
        /// </summary>
        public interface ISheetContentsHandler
        {
            /// <summary>
            /// A row with the (zero based) row number has started */
            /// </summary>
            public void StartRow(int rowNum);
            /// <summary>
            /// A row with the (zero based) row number has ended */
            /// </summary>
            public void EndRow(int rowNum);
            /// <summary>
            /// A cell, with the given formatted value (may be null),
            ///  and possibly a comment (may be null), was encountered */
            /// </summary>
            public void Cell(String cellReference, String formattedValue, XSSFComment comment);
            /// <summary>
            /// A header or footer has been encountered */
            /// </summary>
            public void HeaderFooter(String text, bool IsHeader, String tagName);
        }
    }
}


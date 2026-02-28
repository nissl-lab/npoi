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
using System.IO;
using System.Text;

namespace NPOI.XSSF.Binary
{
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF.EventUserModel;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBSheetHandler : XSSFBParser
    {
        private static  int CHECK_ALL_ROWS = -1;

        private  XSSFBSharedStringsTable _stringsTable;
        private  XSSFSheetXMLHandler.ISheetContentsHandler _handler;
        private  XSSFBStylesTable _styles;
        private  XSSFBCommentsTable _comments;
        private  DataFormatter _dataFormatter;
        private  bool _formulasNotResults;//TODO: implement this

        private int _lastEndedRow = -1;
        private int _lastStartedRow = -1;
        private int _currentRow = 0;
        private byte[] _rkBuffer = new byte[8];
        private XSSFBCellRange _hyperlinkCellRange = null;
        private StringBuilder _xlWideStringBuilder = new StringBuilder();

        private  XSSFBCellHeader _cellBuffer = new XSSFBCellHeader();
        public XSSFBSheetHandler(Stream is1,
                                 XSSFBStylesTable styles,
                                 XSSFBCommentsTable comments,
                                 XSSFBSharedStringsTable strings,
                                 XSSFSheetXMLHandler.ISheetContentsHandler sheetContentsHandler,
                                 DataFormatter dataFormatter,
                                 bool formulasNotResults)
            : base(is1)
        {
            this._styles = styles;
            this._comments = comments;
            this._stringsTable = strings;
            this._handler = sheetContentsHandler;
            this._dataFormatter = dataFormatter;
            this._formulasNotResults = formulasNotResults;
        }
        public override void HandleRecord(int id, byte[] data)
        {
            XSSFBRecordType type = XSSFBRecordTypeClass.Lookup(id);

            switch(type)
            {
                case XSSFBRecordType.BrtRowHdr:
                    int rw = XSSFBUtils.CastToInt(LittleEndian.GetUInt(data, 0));
                    if(rw > 0x00100000)
                    {//could make sure this is larger than currentRow, according to spec?
                        throw new XSSFBParseException("Row number beyond allowable range: "+rw);
                    }
                    _currentRow = rw;
                    CheckMissedComments(_currentRow);
                    StartRow(_currentRow);
                    break;
                case XSSFBRecordType.BrtCellIsst:
                    HandleBrtCellIsst(data);
                    break;
                case XSSFBRecordType.BrtCellSt: //TODO: needs test
                    HandleCellSt(data);
                    break;
                case XSSFBRecordType.BrtCellRk:
                    HandleCellRk(data);
                    break;
                case XSSFBRecordType.BrtCellReal:
                    HandleCellReal(data);
                    break;
                case XSSFBRecordType.BrtCellBool:
                    HandleBoolean(data);
                    break;
                case XSSFBRecordType.BrtCellError:
                    HandleCellError(data);
                    break;
                case XSSFBRecordType.BrtCellBlank:
                    BeforeCellValue(data);//read cell info and check for missing comments
                    break;
                case XSSFBRecordType.BrtFmlaString:
                    HandleFmlaString(data);
                    break;
                case XSSFBRecordType.BrtFmlaNum:
                    HandleFmlaNum(data);
                    break;
                case XSSFBRecordType.BrtFmlaError:
                    HandleFmlaError(data);
                    break;
                //TODO: All the PCDI and PCDIA
                case XSSFBRecordType.BrtEndSheetData:
                    CheckMissedComments(CHECK_ALL_ROWS);
                    EndRow(_lastStartedRow);
                    break;
                case XSSFBRecordType.BrtBeginHeaderFooter:
                    HandleHeaderFooter(data);
                    break;
            }
        }


        private void BeforeCellValue(byte[] data)
        {
            XSSFBCellHeader.Parse(data, 0, _currentRow, _cellBuffer);
            CheckMissedComments(_currentRow, _cellBuffer.ColNum);
        }

        private void HandleCellValue(String formattedValue)
        {
            CellAddress cellAddress = new CellAddress(_currentRow, _cellBuffer.ColNum);
            XSSFBComment comment = null;
            if(_comments != null)
            {
                comment = _comments.Get(cellAddress);
            }
            _handler.Cell(cellAddress.FormatAsString(), formattedValue, comment);
        }

        private void HandleFmlaNum(byte[] data)
        {
            BeforeCellValue(data);
            //xNum
            double val = LittleEndian.GetDouble(data, XSSFBCellHeader.Length);
            HandleCellValue(FormatVal(val, _cellBuffer.StyleIdx));
        }

        private void HandleCellSt(byte[] data)
        {
            BeforeCellValue(data);
            _xlWideStringBuilder.Length = (0);
            XSSFBUtils.ReadXLWideString(data, XSSFBCellHeader.Length, _xlWideStringBuilder);
            HandleCellValue(_xlWideStringBuilder.ToString());
        }

        private void HandleFmlaString(byte[] data)
        {
            BeforeCellValue(data);
            _xlWideStringBuilder.Length = (0);
            XSSFBUtils.ReadXLWideString(data, XSSFBCellHeader.Length, _xlWideStringBuilder);
            HandleCellValue(_xlWideStringBuilder.ToString());
        }

        private void HandleCellError(byte[] data)
        {
            BeforeCellValue(data);
            //TODO, read byte to figure out the type of error
            HandleCellValue("ERROR");
        }

        private void HandleFmlaError(byte[] data)
        {
            BeforeCellValue(data);
            //TODO, read byte to figure out the type of error
            HandleCellValue("ERROR");
        }

        private void HandleBoolean(byte[] data)
        {
            BeforeCellValue(data);
            String formattedVal = (data[XSSFBCellHeader.Length] == 1) ? "TRUE" : "FALSE";
            HandleCellValue(formattedVal);
        }

        private void HandleCellReal(byte[] data)
        {
            BeforeCellValue(data);
            //xNum
            double val = LittleEndian.GetDouble(data, XSSFBCellHeader.Length);
            HandleCellValue(FormatVal(val, _cellBuffer.StyleIdx));
        }

        private void HandleCellRk(byte[] data)
        {
            BeforeCellValue(data);
            double val = RkNumber(data, XSSFBCellHeader.Length);
            HandleCellValue(FormatVal(val, _cellBuffer.StyleIdx));
        }

        private String FormatVal(double val, int styleIdx)
        {
            String formatString = _styles.GetNumberFormatString(styleIdx);
            short styleIndex = _styles.GetNumberFormatIndex(styleIdx);
            //for now, if formatString is null, silently punt
            //and use "General".  Not the best behavior,
            //but we're doing it now in the streaming and non-streaming
            //extractors for xlsx.  See BUG-61053
            if(formatString == null)
            {
                formatString = BuiltinFormats.GetBuiltinFormat(0);
                styleIndex = 0;
            }
            return _dataFormatter.FormatRawCellContents(val, styleIndex, formatString);
        }

        private void HandleBrtCellIsst(byte[] data)
        {
            BeforeCellValue(data);
            int idx = XSSFBUtils.CastToInt(LittleEndian.GetUInt(data, XSSFBCellHeader.Length));
            XSSFRichTextString rtss = new XSSFRichTextString(_stringsTable.GetEntryAt(idx));
            HandleCellValue(rtss.String);
        }


        private void HandleHeaderFooter(byte[] data)
        {
            XSSFBHeaderFooters headerFooter = XSSFBHeaderFooters.parse(data);
            OutputHeaderFooter(headerFooter.Header);
            OutputHeaderFooter(headerFooter.Footer);
            OutputHeaderFooter(headerFooter.HeaderEven);
            OutputHeaderFooter(headerFooter.FooterEven);
            OutputHeaderFooter(headerFooter.HeaderFirst);
            OutputHeaderFooter(headerFooter.FooterFirst);
        }

        private void OutputHeaderFooter(XSSFBHeaderFooter headerFooter)
        {
            String text = headerFooter.String;
            if(text != null && text.Trim().Length > 0)
            {
                _handler.HeaderFooter(text, headerFooter.IsHeader, headerFooter.HeaderFooterTypeLabel);
            }
        }


        //at start of next cell or end of row, return the cellAddress if it equals currentRow and col
        private void CheckMissedComments(int currentRow, int colNum)
        {
            if(_comments == null)
            {
                return;
            }
            Queue<CellAddress> queue = _comments.Addresses;
            while(queue.Count > 0)
            {
                CellAddress cellAddress = queue.Peek();
                if(cellAddress.Row == currentRow && cellAddress.Column < colNum)
                {
                    cellAddress = queue.Dequeue();
                    DumpEmptyCellComment(cellAddress, _comments.Get(cellAddress));
                }
                else if(cellAddress.Row == currentRow && cellAddress.Column == colNum)
                {
                    queue.Dequeue();
                    return;
                }
                else if(cellAddress.Row == currentRow && cellAddress.Column > colNum)
                {
                    return;
                }
                else if(cellAddress.Row > currentRow)
                {
                    return;
                }
            }
        }

        //check for anything from rows before
        private void CheckMissedComments(int currentRow)
        {
            if(_comments == null)
            {
                return;
            }
            Queue<CellAddress> queue = _comments.Addresses;
            int lastInterpolatedRow = -1;
            while(queue.Count > 0)
            {
                CellAddress cellAddress = queue.Peek();
                if(currentRow == CHECK_ALL_ROWS || cellAddress.Row < currentRow)
                {
                    cellAddress = queue.Dequeue();
                    if(cellAddress.Row != lastInterpolatedRow)
                    {
                        StartRow(cellAddress.Row);
                    }
                    DumpEmptyCellComment(cellAddress, _comments.Get(cellAddress));
                    lastInterpolatedRow = cellAddress.Row;
                }
                else
                {
                    break;
                }
            }

        }

        private void StartRow(int row)
        {
            if(row == _lastStartedRow)
            {
                return;
            }

            if(_lastStartedRow != _lastEndedRow)
            {
                EndRow(_lastStartedRow);
            }
            _handler.StartRow(row);
            _lastStartedRow = row;
        }

        private void EndRow(int row)
        {
            if(_lastEndedRow == row)
            {
                return;
            }
            _handler.EndRow(row);
            _lastEndedRow = row;
        }

        private void DumpEmptyCellComment(CellAddress cellAddress, XSSFBComment comment)
        {
            _handler.Cell(cellAddress.FormatAsString(), null, comment);
        }

        private double RkNumber(byte[] data, int offset)
        {
            //see 2.5.122 for this abomination
            byte b0 = data[offset];

            //String s = Int32.ToString(b0, 2);
            bool numDivBy100 = ((b0 & 1) == 1); // else as is
            bool floatingPoint = ((b0 >> 1 & 1) == 0); // else signed integer

            //unset highest 2 bits
            b0 &= unchecked((byte) ~1);
            b0 &= unchecked((byte) ~(1<<1));

            _rkBuffer[4] = b0;
            for(int i = 1; i < 4; i++)
            {
                _rkBuffer[i+4] = data[offset+i];
            }
            double d = 0.0;
            if(floatingPoint)
            {
                d = LittleEndian.GetDouble(_rkBuffer);
            }
            else
            {
                d = LittleEndian.GetInt(_rkBuffer);
            }
            d = (numDivBy100) ? d/100 : d;
            return d;
        }

        /// <summary>
        /// You need to implement this to handle the results
        ///  of the sheet parsing.
        /// </summary>
        public interface ISheetContentsHandler : XSSFSheetXMLHandler.ISheetContentsHandler
        {
            /// <summary>
            /// A cell, with the given formatted value (may be null),
            /// a url (may be null), a toolTip (may be null)
            ///  and possibly a comment (may be null), was encountered */
            /// </summary>
            void HyperlinkCell(String cellReference, String formattedValue, String url, String toolTip, XSSFComment comment);
        }
    }
}


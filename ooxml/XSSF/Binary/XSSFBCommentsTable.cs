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

namespace NPOI.XSSF.Binary
{

    using NPOI.SS.Util;
    using NPOI.Util;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBCommentsTable : XSSFBParser
    {

        private Dictionary<CellAddress, XSSFBComment> comments = new Dictionary<CellAddress, XSSFBComment>();
        private Queue<CellAddress> commentAddresses = new Queue<CellAddress>();
        private List<String> authors = new List<String>();

        //these are all used only during parsing, and they are mutable!
        private int authorId = -1;
        private CellAddress cellAddress = null;
        private XSSFBCellRange cellRange = null;
        private String comment = null;
        private StringBuilder authorBuffer = new StringBuilder();


        public XSSFBCommentsTable(Stream is1)
        : base(is1)
        {
            Parse();
            foreach (var key in comments.Keys)
            {
                commentAddresses.Enqueue(key);
            }
            
        }
        public override void HandleRecord(int id, byte[] data)
        {

            XSSFBRecordType recordType = XSSFBRecordTypeClass.Lookup(id);
            switch(recordType)
            {
                case XSSFBRecordType.BrtBeginComment:
                    int offset = 0;
                    authorId = XSSFBUtils.CastToInt(LittleEndian.GetUInt(data));
                    offset += LittleEndian.INT_SIZE;
                    cellRange = XSSFBCellRange.parse(data, offset, cellRange);
                    offset+= XSSFBCellRange.Length;
                    //for strict parsing; confirm that firstRow==lastRow and firstCol==colLats (2.4.28)
                    cellAddress = new CellAddress(cellRange.firstRow, cellRange.firstCol);
                    break;
                case XSSFBRecordType.BrtCommentText:
                    XSSFBRichStr xssfbRichStr = XSSFBRichStr.Build(data, 0);
                    comment = xssfbRichStr.String;
                    break;
                case XSSFBRecordType.BrtEndComment:
                    comments[cellAddress] = new XSSFBComment(cellAddress, authors[authorId], comment);
                    authorId = -1;
                    cellAddress = null;
                    break;
                case XSSFBRecordType.BrtCommentAuthor:
                    authorBuffer.Length = 0;
                    XSSFBUtils.ReadXLWideString(data, 0, authorBuffer);
                    authors.Add(authorBuffer.ToString());
                    break;
            }
        }


        public Queue<CellAddress> Addresses => commentAddresses;

        public XSSFBComment Get(CellAddress cellAddress)
        {
            if(cellAddress == null)
            {
                return null;
            }
            return comments.TryGetValue(cellAddress, out XSSFBComment value) ? value : null;
        }
    }
}


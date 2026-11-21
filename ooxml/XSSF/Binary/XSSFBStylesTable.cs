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

    using NPOI;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    /// <summary>
    /// This is a very thin shim to gather number formats from styles.bin
    /// files.
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBStylesTable : XSSFBParser
    {

        private  SortedDictionary<short, String> numberFormats = new SortedDictionary<short,String>();
        private  List<short> styleIds = new List<short>();

        private bool inCellXFS = false;
        private bool inFmts = false;
        public XSSFBStylesTable(Stream is1)
        : base(is1)
        {
            Parse();
        }

        public String GetNumberFormatString(int idx)
        {
            short numberFormatIdx = GetNumberFormatIndex(idx);
            if(numberFormats.TryGetValue(numberFormatIdx, out string value))
            {
                return value;
            }

            return BuiltinFormats.GetBuiltinFormat(numberFormatIdx);
        }

        public short GetNumberFormatIndex(int idx)
        {
            return styleIds[idx];
        }
        public override void HandleRecord(int recordType, byte[] data)
        {

            XSSFBRecordType type = XSSFBRecordTypeClass.Lookup(recordType);
            switch(type)
            {
                case XSSFBRecordType.BrtBeginCellXFs:
                    inCellXFS = true;
                    break;
                case XSSFBRecordType.BrtEndCellXFs:
                    inCellXFS = false;
                    break;
                case XSSFBRecordType.BrtXf:
                    if(inCellXFS)
                    {
                        handleBrtXFInCellXF(data);
                    }
                    break;
                case XSSFBRecordType.BrtBeginFmts:
                    inFmts = true;
                    break;
                case XSSFBRecordType.BrtEndFmts:
                    inFmts = false;
                    break;
                case XSSFBRecordType.BrtFmt:
                    if(inFmts)
                    {
                        handleFormat(data);
                    }
                    break;

            }
        }

        private void handleFormat(byte[] data)
        {
            int ifmt = data[0] & 0xFF;
            if(ifmt > short.MaxValue)
            {
                throw new POIXMLException("Format id must be a short");
            }
            StringBuilder sb = new StringBuilder();
            XSSFBUtils.ReadXLWideString(data, 2, sb);
            String fmt = sb.ToString();
            numberFormats[(short) ifmt] = fmt;
        }

        private void handleBrtXFInCellXF(byte[] data)
        {
            int ifmtOffset = 2;
            //int ifmtLength = 2;

            //numFmtId in xml terms
            int ifmt = data[ifmtOffset] & 0xFF;//the second byte is ignored
            styleIds.Add((short) ifmt);
        }
    }
}


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

namespace NPOI.HSSF.Record.PivotTable
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;


    /**
     * SXVIEW - View Definition (0x00B0)<br/>
     * 
     * @author Patrick Cheng
     */
    public class ViewDefinitionRecord : StandardRecord
    {
        public const short sid = 0x00B0;

        private int rwFirst;
        private int rwLast;
        private int colFirst;
        private int colLast;
        private int rwFirstHead;
        private int rwFirstData;
        private int colFirstData;
        private int iCache;
        private int reserved;

        private int sxaxis4Data;
        private int ipos4Data;
        private int cDim;

        private int cDimRw;

        private int cDimCol;
        private int cDimPg;

        private int cDimData;
        private int cRw;
        private int cCol;
        private int grbit;
        private int itblAutoFmt;

        private String dataField;
        private String name;


        public ViewDefinitionRecord(RecordInputStream in1)
        {
            rwFirst = in1.ReadUShort();
            rwLast = in1.ReadUShort();
            colFirst = in1.ReadUShort();
            colLast = in1.ReadUShort();
            rwFirstHead = in1.ReadUShort();
            rwFirstData = in1.ReadUShort();
            colFirstData = in1.ReadUShort();
            iCache = in1.ReadUShort();
            reserved = in1.ReadUShort();
            sxaxis4Data = in1.ReadUShort();
            ipos4Data = in1.ReadUShort();
            cDim = in1.ReadUShort();
            cDimRw = in1.ReadUShort();
            cDimCol = in1.ReadUShort();
            cDimPg = in1.ReadUShort();
            cDimData = in1.ReadUShort();
            cRw = in1.ReadUShort();
            cCol = in1.ReadUShort();
            grbit = in1.ReadUShort();
            itblAutoFmt = in1.ReadUShort();
            int cchName = in1.ReadUShort();
            int cchData = in1.ReadUShort();

            name = StringUtil.ReadUnicodeString(in1, cchName);
            dataField = StringUtil.ReadUnicodeString(in1, cchData);
        }


        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(rwFirst);
            out1.WriteShort(rwLast);
            out1.WriteShort(colFirst);
            out1.WriteShort(colLast);
            out1.WriteShort(rwFirstHead);
            out1.WriteShort(rwFirstData);
            out1.WriteShort(colFirstData);
            out1.WriteShort(iCache);
            out1.WriteShort(reserved);
            out1.WriteShort(sxaxis4Data);
            out1.WriteShort(ipos4Data);
            out1.WriteShort(cDim);
            out1.WriteShort(cDimRw);
            out1.WriteShort(cDimCol);
            out1.WriteShort(cDimPg);
            out1.WriteShort(cDimData);
            out1.WriteShort(cRw);
            out1.WriteShort(cCol);
            out1.WriteShort(grbit);
            out1.WriteShort(itblAutoFmt);
            out1.WriteShort(name.Length);
            out1.WriteShort(dataField.Length);

            StringUtil.WriteUnicodeStringFlagAndData(out1, name);
            StringUtil.WriteUnicodeStringFlagAndData(out1, dataField);
        }


        protected override int DataSize
        {
            get
            {
                return 40 + // 20 short fields (rwFirst ... itblAutoFmt)
                    StringUtil.GetEncodedSize(name) + StringUtil.GetEncodedSize(dataField);
            }
        }


        public override short Sid
        {
            get
            {
                return sid;
            }
        }


        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SXVIEW]\n");
            buffer.Append("    .rwFirst      =").Append(HexDump.ShortToHex(rwFirst)).Append('\n');
            buffer.Append("    .rwLast       =").Append(HexDump.ShortToHex(rwLast)).Append('\n');
            buffer.Append("    .colFirst     =").Append(HexDump.ShortToHex(colFirst)).Append('\n');
            buffer.Append("    .colLast      =").Append(HexDump.ShortToHex(colLast)).Append('\n');
            buffer.Append("    .rwFirstHead  =").Append(HexDump.ShortToHex(rwFirstHead)).Append('\n');
            buffer.Append("    .rwFirstData  =").Append(HexDump.ShortToHex(rwFirstData)).Append('\n');
            buffer.Append("    .colFirstData =").Append(HexDump.ShortToHex(colFirstData)).Append('\n');
            buffer.Append("    .iCache       =").Append(HexDump.ShortToHex(iCache)).Append('\n');
            buffer.Append("    .reserved     =").Append(HexDump.ShortToHex(reserved)).Append('\n');
            buffer.Append("    .sxaxis4Data  =").Append(HexDump.ShortToHex(sxaxis4Data)).Append('\n');
            buffer.Append("    .ipos4Data    =").Append(HexDump.ShortToHex(ipos4Data)).Append('\n');
            buffer.Append("    .cDim         =").Append(HexDump.ShortToHex(cDim)).Append('\n');
            buffer.Append("    .cDimRw       =").Append(HexDump.ShortToHex(cDimRw)).Append('\n');
            buffer.Append("    .cDimCol      =").Append(HexDump.ShortToHex(cDimCol)).Append('\n');
            buffer.Append("    .cDimPg       =").Append(HexDump.ShortToHex(cDimPg)).Append('\n');
            buffer.Append("    .cDimData     =").Append(HexDump.ShortToHex(cDimData)).Append('\n');
            buffer.Append("    .cRw          =").Append(HexDump.ShortToHex(cRw)).Append('\n');
            buffer.Append("    .cCol         =").Append(HexDump.ShortToHex(cCol)).Append('\n');
            buffer.Append("    .grbit        =").Append(HexDump.ShortToHex(grbit)).Append('\n');
            buffer.Append("    .itblAutoFmt  =").Append(HexDump.ShortToHex(itblAutoFmt)).Append('\n');
            buffer.Append("    .name         =").Append(name).Append('\n');
            buffer.Append("    .dataField    =").Append(dataField).Append('\n');

            buffer.Append("[/SXVIEW]\n");
            return buffer.ToString();
        }
    }
}
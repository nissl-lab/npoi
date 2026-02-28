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
     * SXDI - Data Item (0x00C5)<br/>
     * 
     * @author Patrick Cheng
     */
    public class DataItemRecord : StandardRecord
    {
        public const short sid = 0x00C5;

        private int isxvdData;
        private int iiftab;
        private int df;
        private int isxvd;
        private int isxvi;
        private int ifmt;
        private String name;

        public DataItemRecord(RecordInputStream in1)
        {
            isxvdData = in1.ReadUShort();
            iiftab = in1.ReadUShort();
            df = in1.ReadUShort();
            isxvd = in1.ReadUShort();
            isxvi = in1.ReadUShort();
            ifmt = in1.ReadUShort();

            name = in1.ReadString();
        }


        public override void Serialize(ILittleEndianOutput out1)
        {

            out1.WriteShort(isxvdData);
            out1.WriteShort(iiftab);
            out1.WriteShort(df);
            out1.WriteShort(isxvd);
            out1.WriteShort(isxvi);
            out1.WriteShort(ifmt);

            StringUtil.WriteUnicodeString(out1, name);
        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2 + 2 + 2 + 2 + StringUtil.GetEncodedSize(name);
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

            buffer.Append("[SXDI]\n");
            buffer.Append("  .isxvdData = ").Append(HexDump.ShortToHex(isxvdData)).Append("\n");
            buffer.Append("  .iiftab = ").Append(HexDump.ShortToHex(iiftab)).Append("\n");
            buffer.Append("  .df = ").Append(HexDump.ShortToHex(df)).Append("\n");
            buffer.Append("  .isxvd = ").Append(HexDump.ShortToHex(isxvd)).Append("\n");
            buffer.Append("  .isxvi = ").Append(HexDump.ShortToHex(isxvi)).Append("\n");
            buffer.Append("  .ifmt = ").Append(HexDump.ShortToHex(ifmt)).Append("\n");
            buffer.Append("[/SXDI]\n");
            return buffer.ToString();
        }
    }
}
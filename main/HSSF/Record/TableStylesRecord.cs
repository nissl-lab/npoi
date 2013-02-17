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

namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.Util;
    using System.Text;


    /**
     * TABLESTYLES (0x088E)<br/>
     * 
     * @author Patrick Cheng
     */
    public class TableStylesRecord : StandardRecord
    {
        public const short sid = 0x088E;

        private int rt;
        private int grbitFrt;
        private byte[] unused = new byte[8];
        private int cts;

        private String rgchDefListStyle;
        private String rgchDefPivotStyle;


        public TableStylesRecord(RecordInputStream in1)
        {
            rt = in1.ReadUShort();
            grbitFrt = in1.ReadUShort();
            in1.ReadFully(unused);
            cts = in1.ReadInt();
            int cchDefListStyle = in1.ReadUShort();
            int cchDefPivotStyle = in1.ReadUShort();

            rgchDefListStyle = in1.ReadUnicodeLEString(cchDefListStyle);
            rgchDefPivotStyle = in1.ReadUnicodeLEString(cchDefPivotStyle);
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(rt);
            out1.WriteShort(grbitFrt);
            out1.Write(unused);
            out1.WriteInt(cts);

            out1.WriteShort(rgchDefListStyle.Length);
            out1.WriteShort(rgchDefPivotStyle.Length);

            StringUtil.PutUnicodeLE(rgchDefListStyle, out1);
            StringUtil.PutUnicodeLE(rgchDefPivotStyle, out1);
        }


        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 8 + 4 + 2 + 2
                    + (2 * rgchDefListStyle.Length) + (2 * rgchDefPivotStyle.Length);
            }
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[TABLESTYLES]\n");
            buffer.Append("    .rt      =").Append(HexDump.ShortToHex(rt)).Append('\n');
            buffer.Append("    .grbitFrt=").Append(HexDump.ShortToHex(grbitFrt)).Append('\n');
            buffer.Append("    .unused  =").Append(HexDump.ToHex(unused)).Append('\n');
            buffer.Append("    .cts=").Append(HexDump.IntToHex(cts)).Append('\n');
            buffer.Append("    .rgchDefListStyle=").Append(rgchDefListStyle).Append('\n');
            buffer.Append("    .rgchDefPivotStyle=").Append(rgchDefPivotStyle).Append('\n');

            buffer.Append("[/TABLESTYLES]\n");
            return buffer.ToString();
        }
    }
}




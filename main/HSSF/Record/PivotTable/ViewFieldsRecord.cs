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
     * SXVD - View Fields (0x00B1)<br/>
     * 
     * @author Patrick Cheng
     */
    public class ViewFieldsRecord : StandardRecord
    {
        public const short sid = 0x00B1;

		/**
         * values for the {@link ViewFieldsRecord#sxaxis} field
         */
		private enum Axis
		{
			NoAxis = 0,
			Row = 1,
			Column = 2,
			Page = 4,
			Data = 8
		}

        /** the value of the <c>cchName</c> field when the name is not present */
        private const int STRING_NOT_PRESENT_LEN = 0xFFFF;
        /** 5 shorts */
	    private const int BASE_SIZE = 10;
        private int sxaxis;
        private int cSub;
        private int grbitSub;
        private int cItm;

        private String _name = null;

        public ViewFieldsRecord(RecordInputStream in1)
        {
            sxaxis = in1.ReadShort();
            cSub = in1.ReadShort();
            grbitSub = in1.ReadShort();
            cItm = in1.ReadShort();

            int cchName = in1.ReadUShort();
            if (cchName != STRING_NOT_PRESENT_LEN)
            {
                int flag = in1.ReadByte();
                if ((flag & 0x01) != 0)
                {
                    _name = in1.ReadUnicodeLEString(cchName);
                }
                else
                {
                    _name = in1.ReadCompressedUnicode(cchName);
                }
            }
        }


        public override void Serialize(ILittleEndianOutput out1)
        {

            out1.WriteShort(sxaxis);
            out1.WriteShort(cSub);
            out1.WriteShort(grbitSub);
            out1.WriteShort(cItm);

            if (_name != null)
            {
                StringUtil.WriteUnicodeString(out1, _name);
            }
            else
            {
                out1.WriteShort(STRING_NOT_PRESENT_LEN);
            }
        }


        protected override int DataSize
        {
            get
            {
                if (_name == null)
                {
                    return BASE_SIZE;
                }
                return BASE_SIZE + 1 // unicode flag 
                        + _name.Length * (StringUtil.HasMultibyte(_name) ? 2 : 1);
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
            buffer.Append("[SXVD]\n");
            buffer.Append("    .sxaxis    = ").Append(HexDump.ShortToHex(sxaxis)).Append('\n');
            buffer.Append("    .cSub      = ").Append(HexDump.ShortToHex(cSub)).Append('\n');
            buffer.Append("    .grbitSub  = ").Append(HexDump.ShortToHex(grbitSub)).Append('\n');
            buffer.Append("    .cItm      = ").Append(HexDump.ShortToHex(cItm)).Append('\n');
            buffer.Append("    .name      = ").Append(_name).Append('\n');

            buffer.Append("[/SXVD]\n");
            return buffer.ToString();
        }
    }
}
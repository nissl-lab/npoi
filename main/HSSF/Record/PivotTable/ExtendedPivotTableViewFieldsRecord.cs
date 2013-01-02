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
     * SXVDEX - Extended PivotTable View Fields (0x0100)<br/>
     * 
     * @author Patrick Cheng
     */
    public class ExtendedPivotTableViewFieldsRecord : StandardRecord
    {
        public const short sid = 0x0100;

        /** the value of the <c>cchSubName</c> field when the subName is not present */
        private const int STRING_NOT_PRESENT_LEN = 0xFFFF;

        private int grbit1;
        private int grbit2;
        private int citmShow;
        private int isxdiSort;
        private int isxdiShow;
        private int reserved1;
        private int reserved2;
        private String subName;

        public ExtendedPivotTableViewFieldsRecord(RecordInputStream in1)
        {

            grbit1 = in1.ReadInt();
            grbit2 = in1.ReadUByte();
            citmShow = in1.ReadUByte();
            isxdiSort = in1.ReadUShort();
            isxdiShow = in1.ReadUShort();
            // This record seems to have different valid encodings
		    switch (in1.Remaining) {
			    case 0:
				    // as per "Microsoft Excel Developer's Kit" book
				    // older version of SXVDEX - doesn't seem to have a sub-total name
				    reserved1 = 0;
				    reserved2 = 0;
				    subName = null;
				    return;
			    case 10:
				    // as per "MICROSOFT OFFICE EXCEL 97-2007 BINARY FILE FORMAT SPECIFICATION" pdf
				    break;
			    default:
				    throw new RecordFormatException("Unexpected remaining size (" + in1.Remaining + ")");
		    }
            int cchSubName = in1.ReadUShort();
            reserved1 = in1.ReadInt();
            reserved2 = in1.ReadInt();
            if (cchSubName != STRING_NOT_PRESENT_LEN)
            {
                subName = in1.ReadUnicodeLEString(cchSubName);
            }
        }


        public override void Serialize(ILittleEndianOutput out1)
        {

            out1.WriteInt(grbit1);
            out1.WriteByte(grbit2);
            out1.WriteByte(citmShow);
            out1.WriteShort(isxdiSort);
            out1.WriteShort(isxdiShow);

            if (subName == null)
            {
                out1.WriteShort(STRING_NOT_PRESENT_LEN);
            }
            else
            {
                out1.WriteShort(subName.Length);
            }

            out1.WriteInt(reserved1);
            out1.WriteInt(reserved2);
            if (subName != null)
            {
                StringUtil.PutUnicodeLE(subName, out1);
            }

        }


        protected override int DataSize
        {
            get
            {
                return 4 + 1 + 1 + 2 + 2 + 2 + 4 + 4 +
                            (subName == null ? 0 : (2 * subName.Length)); // in unicode
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

            buffer.Append("[SXVDEX]\n");

            buffer.Append("    .grbit1 =").Append(HexDump.IntToHex(grbit1)).Append("\n");
            buffer.Append("    .grbit2 =").Append(HexDump.ByteToHex(grbit2)).Append("\n");
            buffer.Append("    .citmShow =").Append(HexDump.ByteToHex(citmShow)).Append("\n");
            buffer.Append("    .isxdiSort =").Append(HexDump.ShortToHex(isxdiSort)).Append("\n");
            buffer.Append("    .isxdiShow =").Append(HexDump.ShortToHex(isxdiShow)).Append("\n");
            buffer.Append("    .subName =").Append(subName).Append("\n");
            buffer.Append("[/SXVDEX]\n");
            return buffer.ToString();
        }
    }
}
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
    using System.Text;
    using NPOI.Util;

    /**
     * Title:        Bound Sheet Record (aka BundleSheet) (0x0085) for BIFF 5<br/>
     * Description:  Defines a sheet within a workbook.  Basically stores the sheet name
     *               and tells where the Beginning of file record is within the HSSF
     *               file.
     */
    public class OldSheetRecord
    {
        public const short sid = 0x0085;

        private int field_1_position_of_BOF;
        private int field_2_visibility;
        private int field_3_type;
        private byte[] field_5_sheetname;
        private CodepageRecord codepage;

        public OldSheetRecord(RecordInputStream in1)
        {
            field_1_position_of_BOF = in1.ReadInt();
            field_2_visibility = in1.ReadUByte();
            field_3_type = in1.ReadUByte();
            int field_4_sheetname_length = in1.ReadUByte();
            field_5_sheetname = new byte[field_4_sheetname_length];
            in1.Read(field_5_sheetname, 0, field_4_sheetname_length);
        }

        public void SetCodePage(CodepageRecord codepage)
        {
            this.codepage = codepage;
        }

        public short Sid
        {
            get
            {
                return sid;
            }
        }

        /**
         * Get the offset in bytes of the Beginning of File Marker within the HSSF Stream part of the POIFS file
         *
         * @return offset in bytes
         */
        public int PositionOfBof
        {
            get
            {
                return field_1_position_of_BOF;
            }
        }

        /**
         * Get the sheetname for this sheet.  (this appears in the tabs at the bottom)
         * @return sheetname the name of the sheet
         */
        public String Sheetname
        {
            get
            {
                return OldStringRecord.GetString(field_5_sheetname, codepage);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[BOUNDSHEET]\n");
            buffer.Append("    .bof        = ").Append(HexDump.IntToHex(PositionOfBof)).Append("\n");
            buffer.Append("    .visibility = ").Append(HexDump.ShortToHex(field_2_visibility)).Append("\n");
            buffer.Append("    .type       = ").Append(HexDump.ByteToHex(field_3_type)).Append("\n");
            buffer.Append("    .sheetname  = ").Append(Sheetname).Append("\n");
            buffer.Append("[/BOUNDSHEET]\n");
            return buffer.ToString();
        }
    }

}
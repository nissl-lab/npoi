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
     * Biff2 - Biff 4 Label Record (0x0004 / 0x0204) - read only support for 
     *  strings stored directly in the cell, from the older file formats that
     *  didn't use {@link LabelSSTRecord}
     */
    public class OldLabelRecord : OldCellRecord
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(OldLabelRecord));

        public const short biff2_sid = 0x0004;
        public const short biff345_sid = 0x0204;

        private short field_4_string_len;
        private byte[] field_5_bytes;
        private CodepageRecord codepage;

        /**
         * @param in the RecordInputstream to read the record from
         */
        public OldLabelRecord(RecordInputStream in1)
            : base(in1, in1.Sid == biff2_sid)
        {
            if (IsBiff2)
            {
                field_4_string_len = (short)in1.ReadUByte();
            }
            else
            {
                field_4_string_len = in1.ReadShort();
            }

            // Can only decode properly later when you know the codepage
            field_5_bytes = new byte[field_4_string_len];
            in1.Read(field_5_bytes, 0, field_4_string_len);

            if (in1.Remaining > 0)
            {
                logger.Log(POILogger.INFO,
                        "LabelRecord data remains: " + in1.Remaining +
                        " : " + HexDump.ToHex(in1.ReadRemainder())
                        );
            }
        }

        public void SetCodePage(CodepageRecord codepage)
        {
            this.codepage = codepage;
        }

        /**
         * Get the number of characters this string Contains
         * @return number of characters
         */
        public short StringLength
        {
            get
            {
                return field_4_string_len;
            }
        }

        /**
         * Get the String of the cell
         */
        public String Value
        {
            get
            {
                return OldStringRecord.GetString(field_5_bytes, codepage);
            }
        }

        /**
         * Not supported
         */
        public int Serialize(int offset, byte[] data)
        {
            throw new RecordFormatException("Old Label Records are supported READ ONLY");
        }
        public int RecordSize
        {
            get
            {
                throw new RecordFormatException("Old Label Records are supported READ ONLY");
            }
        }

        protected override void AppendValueText(StringBuilder sb)
        {
            sb.Append("    .string_len= ").Append(HexDump.ShortToHex(field_4_string_len)).Append("\n");
            sb.Append("    .value       = ").Append(Value).Append("\n");
        }

        protected override String RecordName
        {
            get
            {
                return "OLD LABEL";
            }
        }
    }

}
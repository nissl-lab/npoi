
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;

    /**
     * Title:        Format Record
     * Description:  describes a number format -- those goofy strings like $(#,###)
     *
     * REFERENCE:  PG 317 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Shawn M. Laubach (slaubach at apache dot org)  
     * @version 2.0-pre
     */

    public class FormatRecord
       : StandardRecord
    {
        public const short sid = 0x41e;
        private int field_1_index_code;

        //private short field_3_unicode_len;      // Unicode string Length
        //private bool field_3_unicode_flag;     // it Is not Undocumented - it Is Unicode flag
        private bool field_3_hasMultibyte;
        private String field_4_formatstring;

        public FormatRecord(int indexCode, String fs)
        {
            field_1_index_code = indexCode;
            field_4_formatstring = fs;
            field_3_hasMultibyte = StringUtil.HasMultibyte(fs);
        }

        /**
         * Constructs a Format record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public FormatRecord(RecordInputStream in1)
        {
            field_1_index_code = in1.ReadShort();
            int field_3_unicode_len = in1.ReadShort();
            field_3_hasMultibyte = (in1.ReadByte() & (byte)0x01) != 0;

            if (field_3_hasMultibyte)
            {
                // Unicode
                field_4_formatstring = in1.ReadUnicodeLEString(field_3_unicode_len);
            }
            else
            {
                // not Unicode
                field_4_formatstring = in1.ReadCompressedUnicode(field_3_unicode_len);
            }
        }
        
        /**
         * Get the format index code (for built in formats)
         *
         * @return the format index code
         * @see org.apache.poi.hssf.model.Workbook
         */

        public int IndexCode
        {
            get
            {
                return field_1_index_code;
            }
        }

        /**
         * Get the format string
         *
         * @return the format string
         */

        public String FormatString
        {
            get
            {
                return field_4_formatstring;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FORMAT]\n");
            buffer.Append("    .indexcode       = ").Append(HexDump.ShortToHex(IndexCode)).Append("\n");
            buffer.Append("    .isUnicode       = ").Append(field_3_hasMultibyte).Append("\n");
            buffer.Append("    .formatstring    = ").Append(FormatString).Append("\n");
            buffer.Append("[/FORMAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            String formatString = FormatString;
            out1.WriteShort(IndexCode);
            out1.WriteShort(formatString.Length);
            out1.WriteByte(field_3_hasMultibyte ? 0x01 : 0x00);

            if (field_3_hasMultibyte)
            {
                StringUtil.PutUnicodeLE(formatString, out1);
            }
            else
            {
                StringUtil.PutCompressedUnicode(formatString, out1);
            }
        }
        protected override int DataSize
        {
            get
            {
                return 5 // 2 shorts + 1 byte
                    + FormatString.Length * (field_3_hasMultibyte ? 2 : 1);
            }
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override Object Clone()
        {
            // immutable
            return this;
        }
    }
}

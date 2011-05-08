
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
       : Record
    {
        public const short sid = 0x41e;
        private short field_1_index_code;

        private short field_3_unicode_len;      // Unicode string Length
        private bool field_3_unicode_flag;     // it Is not Undocumented - it Is Unicode flag
        private String field_4_formatstring;

        public FormatRecord()
        {
        }

        /**
         * Constructs a Format record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public FormatRecord(RecordInputStream in1)
        {
            field_1_index_code = in1.ReadShort();
            field_3_unicode_len = in1.ReadShort();
            field_3_unicode_flag = (in1.ReadByte() & (byte)0x01) != 0;

            if (field_3_unicode_flag)
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
         * Set the format index code (for built in formats)
         *
         * @param index  the format index code
         * @see org.apache.poi.hssf.model.Workbook
         */

        public void SetIndexCode(short index)
        {
            field_1_index_code = index;
        }

        /**
         * Set the format string Length
         *
         * @param len  the Length of the format string
         * @see #SetFormatString(String)
         */

        public void SetFormatStringLength(byte len)
        {

            field_3_unicode_len = len;
        }

        /**
         * Set whether the string Is Unicode
         *
         * @param Unicode flag for whether string Is Unicode
         */

        public void SetUnicodeFlag(bool Unicode)
        {
            field_3_unicode_flag = Unicode;
        }

        /**
         * Set the format string
         *
         * @param fs  the format string
         * @see #SetFormatStringLength(byte)
         */

        public void SetFormatString(String fs)
        {
            field_4_formatstring = fs;
            SetUnicodeFlag(StringUtil.HasMultibyte(fs));
        }

        /**
         * Get the format index code (for built in formats)
         *
         * @return the format index code
         * @see org.apache.poi.hssf.model.Workbook
         */

        public short GetIndexCode()
        {
            return field_1_index_code;
        }

        /**
         * Get the format string Length
         *
         * @return the Length of the format string
         * @see #GetFormatString()
         */

        /* public short GetFormatStringLength
         {
             return field_3_unicode_flag ? field_3_unicode_len : field_2_formatstring_len;
         }*/

        /**
         * Get whether the string Is Unicode
         *
         * @return flag for whether string Is Unicode
         */

        public bool GetUnicodeFlag()
        {
            return field_3_unicode_flag;
        }

        /**
         * Get the format string
         *
         * @return the format string
         */

        public String GetFormatString()
        {
            return field_4_formatstring;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FORMAT]\n");
            buffer.Append("    .indexcode       = ")
                .Append(StringUtil.ToHexString(GetIndexCode())).Append("\n");
            /*
            buffer.Append("    .formatstringlen = ")
                .Append(StringUtil.ToHexString(GetFormatStringLength))
                .Append("\n");
            */
            buffer.Append("    .unicode Length  = ")
                .Append(StringUtil.ToHexString(field_3_unicode_len)).Append("\n");
            buffer.Append("    .IsUnicode       = ")
                .Append(field_3_unicode_flag).Append("\n");
            buffer.Append("    .formatstring    = ").Append(GetFormatString())
                .Append("\n");
            buffer.Append("[/FORMAT]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset, (short)(2 + 2 + 1 + ((field_3_unicode_flag)
                                                                      ? 2 * field_3_unicode_len
                                                                      : field_3_unicode_len)));
            // index + len + flag + format string Length
            LittleEndian.PutShort(data, 4 + offset, GetIndexCode());
            LittleEndian.PutShort(data, 6 + offset, field_3_unicode_len);
            data[8 + offset] = (byte)((field_3_unicode_flag) ? 0x01 : 0x00);

            if (field_3_unicode_flag)
            {
                // Unicode
                StringUtil.PutUnicodeLE(GetFormatString(), data, 9 + offset);
            }
            else
            {
                // not Unicode
                StringUtil.PutCompressedUnicode(GetFormatString(), data, 9 + offset);
            }

            return RecordSize;
        }

        public override int RecordSize
        {
            get { return 9 + ((field_3_unicode_flag) ? 2 * field_3_unicode_len : field_3_unicode_len); }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}

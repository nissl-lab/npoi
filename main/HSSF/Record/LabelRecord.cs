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
     * Label Record - Read only support for strings stored directly in the cell..  Don't
     * use this (except to Read), use LabelSST instead 
     * REFERENCE:  PG 325 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * 
     * @see org.apache.poi.hssf.record.LabelSSTRecord
     */
    public class LabelRecord : Record, CellValueRecordInterface
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(LabelRecord));
        public const short sid = 0x204;

        private int field_1_row;
        private int field_2_column;
        private short field_3_xf_index;
        private short field_4_string_len;
        private byte field_5_unicode_flag;
        private String field_6_value;

        /** Creates new LabelRecord */

        public LabelRecord()
        {
        }

        /**
         * Constructs an Label record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public LabelRecord(RecordInputStream in1)
        {
            field_1_row = in1.ReadUShort();
            field_2_column = in1.ReadUShort();
            field_3_xf_index = in1.ReadShort();
            field_4_string_len = in1.ReadShort();
            field_5_unicode_flag = (byte)in1.ReadByte();
            if (field_4_string_len > 0)
            {
                if (IsUncompressedUnicode)
                {
                    field_6_value = in1.ReadUnicodeLEString(field_4_string_len);
                }
                else
                {
                    field_6_value = in1.ReadCompressedUnicode(field_4_string_len);
                }
            }
            else
            {
                field_6_value = "";
            }
            if (in1.Remaining > 0)
            {
                logger.Log(POILogger.INFO, "LabelRecord data remains: " +in1.Remaining +
                " : " + HexDump.ToHex(in1.ReadRemainder()));
            }
        }

        /*
         * Read ONLY ACCESS... THIS Is FOR COMPATIBILITY ONLY...USE LABELSST! public
         * void SetRow(short row) { field_1_row = row; }
         * 
         * public void SetColumn(short col) { field_2_column = col; }
         * 
         * public void SetXFIndex(short index) { field_3_xf_index = index; }
         */
        public int Row
        {
            get{return field_1_row;}
            set { throw new NotSupportedException("Use LabelSST instead"); }
        }

        public int Column
        {
            get{return field_2_column;}
            set { throw new NotSupportedException("Use LabelSST instead"); }
        }

        public short XFIndex
        {
            get { return field_3_xf_index; }
            set { throw new NotSupportedException("Use LabelSST instead"); }
        }

        /**
         * Get the number of Chars this string Contains
         * @return number of Chars
         */

        public short StringLength
        {
            get { return field_4_string_len; }
        }

        /**
         * Is this Uncompressed Unicode (16bit)?  Or just 8-bit compressed?
         * @return IsUnicode - True for 16bit- false for 8bit
         */

        public bool IsUncompressedUnicode
        {
            get { return (field_5_unicode_flag & 0x01) != 0; }
        }

        /**
         * Get the value
         *
         * @return the text string
         * @see #GetStringLength
         */

        public String Value
        {
            get { return field_6_value; }
        }

        /**
         * THROWS A RUNTIME EXCEPTION..  USE LABELSSTRecords.  YOU HAVE NO REASON to use LABELRecord!!
         */

        public override int Serialize(int offset, byte [] data)
        {
            throw new RecordFormatException(
                "Label Records are supported Read ONLY...Convert to LabelSST");
        }
        public override int RecordSize
        {
            get
            {
                throw new RecordFormatException("Label Records are supported READ ONLY...convert to LabelSST");
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[LABEL]\n");
            buffer.Append("    .row            = ")
                .Append(StringUtil.ToHexString(Row)).Append("\n");
            buffer.Append("    .column         = ")
                .Append(StringUtil.ToHexString(Column)).Append("\n");
            buffer.Append("    .xfindex        = ")
                .Append(StringUtil.ToHexString(XFIndex)).Append("\n");
            buffer.Append("    .string_len       = ")
                .Append(StringUtil.ToHexString(field_4_string_len)).Append("\n");
            buffer.Append("    .unicode_flag       = ")
                .Append(StringUtil.ToHexString(field_5_unicode_flag)).Append("\n");
            buffer.Append("    .value       = ")
                .Append(Value).Append("\n");
            buffer.Append("[/LABEL]\n");
            return buffer.ToString();
        }


        public override Object Clone()
        {
            LabelRecord rec = new LabelRecord();
            rec.field_1_row = field_1_row;
            rec.field_2_column = field_2_column;
            rec.field_3_xf_index = field_3_xf_index;
            rec.field_4_string_len = field_4_string_len;
            rec.field_5_unicode_flag = field_5_unicode_flag;
            rec.field_6_value = field_6_value;
            return rec;
        }
    }
}

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


namespace NPOI.HSSF.Record.Chart
{

    using System;
    using System.Text;
    using NPOI.Util;

    /**
     * Defines a series name
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    public class SeriesTextRecord
       : StandardRecord
    {
        /** the actual text cannot be longer than 255 characters */
        private const int MAX_LEN = 0xFF;

        public const short sid = 0x100d;
        private short field_1_id;

        private bool is16bit;
        private String field_4_text;



        public SeriesTextRecord()
        {
            field_4_text = "";
            is16bit = false;
        }

        /**
         * Constructs a SeriesText record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SeriesTextRecord(RecordInputStream in1)
        {

            field_1_id = in1.ReadShort();
            int field_2_textLength = (byte)in1.ReadByte();
            is16bit = (in1.ReadUByte() & 0x01) != 0;
            if (is16bit)
            {
                field_4_text = in1.ReadUnicodeLEString(field_2_textLength);
            }
            else
            {
                field_4_text = in1.ReadCompressedUnicode(field_2_textLength);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SERIESTEXT]\n");
            buffer.Append("    .id                   = ")
                .Append("0x").Append(HexDump.ToHex(Id))
                .Append(" (").Append(Id).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .textLength           = ")
                .Append(field_4_text.Length);
            buffer.Append(Environment.NewLine);
            buffer.Append("    .is16bit         = ").Append(is16bit);
            buffer.Append(Environment.NewLine);
            buffer.Append("    .text                 = ")
                .Append(" (").Append(Text).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SERIESTEXT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_id);
            out1.WriteByte(field_4_text.Length);
            if (is16bit)
            {
                // Excel (2007) seems to choose 16bit regardless of whether it is needed
                out1.WriteByte(0x01);
                StringUtil.PutUnicodeLE(field_4_text, out1);
            }
            else
            {
                // Excel can read this OK
                out1.WriteByte(0x00);
                StringUtil.PutCompressedUnicode(field_4_text, out1);
            }
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 1 + 1 + field_4_text.Length * (is16bit ? 2 : 1); }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            SeriesTextRecord rec = new SeriesTextRecord();

            rec.field_1_id = field_1_id;
            rec.is16bit = is16bit;
            rec.field_4_text = field_4_text;
            return rec;
        }




        /**
         * Get the id field for the SeriesText record.
         */
        public short Id
        {
            get
            {
                return field_1_id;
            }
            set { this.field_1_id = value; }
        }

        /**
         * Get the text field for the SeriesText record.
         */
        public String Text
        {
            get
            {
                return field_4_text;
            }
            set
            {
                if (value.Length > MAX_LEN)
                {
                    throw new ArgumentException("Text is too long ("
                            + value.Length + ">" + MAX_LEN + ")");
                }
                field_4_text = value;
                is16bit = StringUtil.HasMultibyte(value);
            }
        }


    }
}
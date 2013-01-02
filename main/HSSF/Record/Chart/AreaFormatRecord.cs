
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
     * The area format record is used to define the colours and patterns for an area.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class AreaFormatRecord
       : StandardRecord
    {
        public const short FILL_PATTERN_NONE = 0x0;
        public const short FILL_PATTERN_SOLID = 0x1;
        public const short FILL_PATTERN_MEDIUM_GRAY = 0x2;
        public const short FILL_PATTERN_DARK_GRAY = 0x3;
        public const short FILL_PATTERN_LIGHT_GRAY = 0x4;
        public const short FILL_PATTERN_HORIZONTAL_STRIPES = 0x5;
        public const short FILL_PATTERN_VERTICAL_STRIPES = 0x6;
        public const short FILL_PATTERN_DOWNWARD_DIAGONAL_STRIPES = 0x7;
        public const short FILL_PATTERN_UPWARD_DIAGONAL_STRIPES = 0x8;
        public const short FILL_PATTERN_GRID = 0x9;
        public const short FILL_PATTERN_TRELLIS = 0xA;
        public const short FILL_PATTERN_LIGHT_HORIZONTAL_STRIPES = 0xB;
        public const short FILL_PATTERN_LIGHT_VERTICAL_STRIPES = 0xC;
        public const short FILL_PATTERN_LIGHTDOWN= 0xD;
        public const short FILL_PATTERN_LIGHTUP = 0xE;
        public const short FILL_PATTERN_LIGHT_GRID = 0xF;
        public const short FILL_PATTERN_LIGHT_TRELLIS = 0x10;
        public const short FILL_PATTERN_GRAYSCALE_1_8 = 0x11;
        public const short FILL_PATTERN_GRAYSCALE_1_16 = 0x12;

        public const short sid = 0x100a;
        private int field_1_foregroundColor;
        private int field_2_backgroundColor;
        private short field_3_pattern;
        private short field_4_formatFlags;
        private BitField automatic = BitFieldFactory.GetInstance(0x1);
        private BitField invert = BitFieldFactory.GetInstance(0x2);
        private short field_5_forecolorIndex;
        private short field_6_backcolorIndex;


        public AreaFormatRecord()
        {

        }

        /**
         * Constructs a AreaFormat record and s its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public AreaFormatRecord(RecordInputStream in1)
        {

            field_1_foregroundColor = in1.ReadInt();
            field_2_backgroundColor = in1.ReadInt();
            field_3_pattern = in1.ReadShort();
            field_4_formatFlags = in1.ReadShort();
            field_5_forecolorIndex = in1.ReadShort();
            field_6_backcolorIndex = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AREAFORMAT]\n");
            buffer.Append("    .foregroundColor      = ")
                .Append("0x").Append(HexDump.ToHex(ForegroundColor))
                .Append(" (").Append(ForegroundColor).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .backgroundColor      = ")
                .Append("0x").Append(HexDump.ToHex(BackgroundColor))
                .Append(" (").Append(BackgroundColor).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .pattern              = ")
                .Append("0x").Append(HexDump.ToHex(Pattern))
                .Append(" (").Append(Pattern).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .formatFlags          = ")
                .Append("0x").Append(HexDump.ToHex(FormatFlags))
                .Append(" (").Append(FormatFlags).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .automatic                = ").Append(IsAutomatic).Append('\n');
            buffer.Append("         .invert                   = ").Append(IsInvert).Append('\n');
            buffer.Append("    .forecolorIndex       = ")
                .Append("0x").Append(HexDump.ToHex(ForecolorIndex))
                .Append(" (").Append(ForecolorIndex).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .backcolorIndex       = ")
                .Append("0x").Append(HexDump.ToHex(BackcolorIndex))
                .Append(" (").Append(BackcolorIndex).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/AREAFORMAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_foregroundColor);
            out1.WriteInt(field_2_backgroundColor);
            out1.WriteShort(field_3_pattern);
            out1.WriteShort(field_4_formatFlags);
            out1.WriteShort(field_5_forecolorIndex);
            out1.WriteShort(field_6_backcolorIndex);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
             get{ return 4 + 4 + 2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
             get{ return sid; }
        }

        public override Object Clone()
        {
            AreaFormatRecord rec = new AreaFormatRecord();

            rec.field_1_foregroundColor = field_1_foregroundColor;
            rec.field_2_backgroundColor = field_2_backgroundColor;
            rec.field_3_pattern = field_3_pattern;
            rec.field_4_formatFlags = field_4_formatFlags;
            rec.field_5_forecolorIndex = field_5_forecolorIndex;
            rec.field_6_backcolorIndex = field_6_backcolorIndex;
            return rec;
        }




        /**
         *  the foreground color field for the AreaFormat record.
         */
        public int ForegroundColor
        {
            get{return field_1_foregroundColor;}
            set{this.field_1_foregroundColor = value;}
        }

        /**
         *  the background color field for the AreaFormat record.
         */
        public int BackgroundColor
        {
            get{return field_2_backgroundColor;}
            set{this.field_2_backgroundColor=value;}
        }

        /**
         *  the pattern field for the AreaFormat record.
         */
        public short Pattern
        {
            get{return field_3_pattern;}
            set{this.field_3_pattern =value;}
        }

        /**
         *  the format flags field for the AreaFormat record.
         */
        public short FormatFlags
        {
            get{return field_4_formatFlags;}
            set{this.field_4_formatFlags =value;}
        }

        /**
         *  the forecolor index field for the AreaFormat record.
         */
        public short ForecolorIndex
        {
            get{return field_5_forecolorIndex;}
            set{this.field_5_forecolorIndex =value;}
        }

        /**
         *  the backcolor index field for the AreaFormat record.
         */
        public short BackcolorIndex
        {
            get{return field_6_backcolorIndex;}
            set{this.field_6_backcolorIndex =value;}
        }

        /**
         * automatic formatting
         * @return  the automatic field value.
         */
        public bool IsAutomatic
        {
            get{return automatic.IsSet(field_4_formatFlags);}
            set{
                field_4_formatFlags = automatic.SetShortBoolean(field_4_formatFlags, value);
            }
        }

        /**
         * swap foreground and background colours when data is negative
         * @return  the invert field value.
         */
        public bool IsInvert
        {
            get{return invert.IsSet(field_4_formatFlags);}
            set{ field_4_formatFlags = invert.SetShortBoolean(field_4_formatFlags, value);}
        }


    }
}





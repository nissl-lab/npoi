
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
     * Describes a line format record.  The line format record controls how a line on a chart appears.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class LineFormatRecord
       : StandardRecord
    {
        public const short sid = 0x1007;
        private int field_1_lineColor;
        private short field_2_linePattern;
        public const short LINE_PATTERN_SOLID = 0;
        public const short LINE_PATTERN_DASH = 1;
        public const short LINE_PATTERN_DOT = 2;
        public const short LINE_PATTERN_DASH_DOT = 3;
        public const short LINE_PATTERN_DASH_DOT_DOT = 4;
        public const short LINE_PATTERN_NONE = 5;
        public const short LINE_PATTERN_DARK_GRAY_PATTERN = 6;
        public const short LINE_PATTERN_MEDIUM_GRAY_PATTERN = 7;
        public const short LINE_PATTERN_LIGHT_GRAY_PATTERN = 8;
        private short field_3_weight;
        public const short WEIGHT_HAIRLINE = -1;
        public const short WEIGHT_NARROW = 0;
        public const short WEIGHT_MEDIUM = 1;
        public const short WEIGHT_WIDE = 2;
        private short field_4_format;
        private BitField auto = BitFieldFactory.GetInstance(0x1);
        private BitField drawTicks = BitFieldFactory.GetInstance(0x4);
        private BitField Unknown = BitFieldFactory.GetInstance(0x8);
        private short field_5_colourPaletteIndex;


        public LineFormatRecord()
        {

        }

        /**
         * Constructs a LineFormat record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public LineFormatRecord(RecordInputStream in1)
        {
            field_1_lineColor = in1.ReadInt();
            field_2_linePattern = in1.ReadShort();
            field_3_weight = in1.ReadShort();
            field_4_format = in1.ReadShort();
            field_5_colourPaletteIndex = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[LINEFORMAT]\n");
            buffer.Append("    .lineColor            = ")
                .Append("0x").Append(HexDump.ToHex(LineColor))
                .Append(" (").Append(LineColor).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .linePattern          = ")
                .Append("0x").Append(HexDump.ToHex(LinePattern))
                .Append(" (").Append(LinePattern).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .weight               = ")
                .Append("0x").Append(HexDump.ToHex(Weight))
                .Append(" (").Append(Weight).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .format               = ")
                .Append("0x").Append(HexDump.ToHex(Format))
                .Append(" (").Append(Format).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .auto                     = ").Append(IsAuto).Append('\n');
            buffer.Append("         .drawTicks                = ").Append(IsDrawTicks).Append('\n');
            buffer.Append("         .unknown                  = ").Append(IsUnknown).Append('\n');
            buffer.Append("    .colourPaletteIndex   = ")
                .Append("0x").Append(HexDump.ToHex(ColourPaletteIndex))
                .Append(" (").Append(ColourPaletteIndex).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/LINEFORMAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_lineColor);
            out1.WriteShort(field_2_linePattern);
            out1.WriteShort(field_3_weight);
            out1.WriteShort(field_4_format);
            out1.WriteShort(field_5_colourPaletteIndex);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return  4 + 2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            LineFormatRecord rec = new LineFormatRecord();

            rec.field_1_lineColor = field_1_lineColor;
            rec.field_2_linePattern = field_2_linePattern;
            rec.field_3_weight = field_3_weight;
            rec.field_4_format = field_4_format;
            rec.field_5_colourPaletteIndex = field_5_colourPaletteIndex;
            return rec;
        }




        /**
         * Get the line color field for the LineFormat record.
         */
        public int LineColor
        {
            get
            {
                return field_1_lineColor;
            }
            set 
            {
                this.field_1_lineColor = value;
            }
        }
        /**
         * Get the line pattern field for the LineFormat record.
         *
         * @return  One of 
         *        LINE_PATTERN_SOLID
         *        LINE_PATTERN_DASH
         *        LINE_PATTERN_DOT
         *        LINE_PATTERN_DASH_DOT
         *        LINE_PATTERN_DASH_DOT_DOT
         *        LINE_PATTERN_NONE
         *        LINE_PATTERN_DARK_GRAY_PATTERN
         *        LINE_PATTERN_MEDIUM_GRAY_PATTERN
         *        LINE_PATTERN_LIGHT_GRAY_PATTERN
         */
        public short LinePattern
        {
            get
            {
                return field_2_linePattern;
            }
            set 
            {
                this.field_2_linePattern = value;
                if (value == LINE_PATTERN_NONE)
                {
                    field_3_weight = (short)-1;
                    field_5_colourPaletteIndex = 0x004D;
                }
            }
        }

        /**
         * Get the weight field for the LineFormat record.
         * specifies the thickness of the line.
         * @return  One of 
         *        WEIGHT_HAIRLINE
         *        WEIGHT_NARROW
         *        WEIGHT_MEDIUM
         *        WEIGHT_WIDE
         */
        public short Weight
        {
            get
            {
                return field_3_weight;
            }
            set 
            {
                this.field_3_weight = value;
            }
        }

        /**
         * Get the format field for the LineFormat record.
         */
        public short Format
        {
            get
            {
                return field_4_format;
            }
            set 
            {
                this.field_4_format = value;
            }
        }

        /**
         * Get the colour palette index field for the LineFormat record.
         */
        public short ColourPaletteIndex
        {
            get
            {
                return field_5_colourPaletteIndex;
            }
            set { this.field_5_colourPaletteIndex = value; }
        }

        /**
         * automatic format
         * @return  the auto field value.
         */
        public bool IsAuto
        {
            get
            {
                return auto.IsSet(field_4_format);
            }
            set 
            {
                field_4_format = auto.SetShortBoolean(field_4_format, value);
            }
        }

        /**
         * draw tick marks
         * @return  the draw ticks field value.
         */
        public bool IsDrawTicks
        {
            get
            {
                return drawTicks.IsSet(field_4_format);
            }
            set 
            {
                field_4_format = drawTicks.SetShortBoolean(field_4_format, value);
            }
        }

        /**
         * book marks this as reserved = 0 but it seems to do something
         * @return  the Unknown field value.
         */
        public bool IsUnknown
        {
            get
            {
                return Unknown.IsSet(field_4_format);
            }
            set 
            {
                field_4_format = Unknown.SetShortBoolean(field_4_format, value);
            }
        }


    }
}



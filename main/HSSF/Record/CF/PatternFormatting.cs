
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
        

/*
 * FontFormatting.java
 *
 * Created on January 22, 2008, 10:05 PM
 */
namespace NPOI.HSSF.Record.CF
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;


    /**
     * Pattern Formatting Block of the Conditional Formatting Rule Record.
     * 
     * @author Dmitriy Kumshayev
     */

    public class PatternFormatting : ICloneable
    {
        /**  No background */
        public const short NO_Fill = 0;
        /**  Solidly Filled */
        public const short SOLID_FOREGROUND = 1;
        /**  Small fine dots */
        public const short FINE_DOTS = 2;
        /**  Wide dots */
        public const short ALT_BARS = 3;
        /**  SParse dots */
        public const short SPARSE_DOTS = 4;
        /**  Thick horizontal bands */
        public const short THICK_HORZ_BANDS = 5;
        /**  Thick vertical bands */
        public const short THICK_VERT_BANDS = 6;
        /**  Thick backward facing diagonals */
        public const short THICK_BACKWARD_DIAG = 7;
        /**  Thick forward facing diagonals */
        public const short THICK_FORWARD_DIAG = 8;
        /**  Large spots */
        public const short BIG_SPOTS = 9;
        /**  Brick-like layout */
        public const short BRICKS = 10;
        /**  Thin horizontal bands */
        public const short THIN_HORZ_BANDS = 11;
        /**  Thin vertical bands */
        public const short THIN_VERT_BANDS = 12;
        /**  Thin backward diagonal */
        public const short THIN_BACKWARD_DIAG = 13;
        /**  Thin forward diagonal */
        public const short THIN_FORWARD_DIAG = 14;
        /**  Squares */
        public const short SQUARES = 15;
        /**  Diamonds */
        public const short DIAMONDS = 16;
        /**  Less Dots */
        public const short LESS_DOTS = 17;
        /**  Least Dots */
        public const short LEAST_DOTS = 18;

        public PatternFormatting()
        {
            field_15_pattern_style = (short)0;
            field_16_pattern_color_indexes = (short)0;
        }

        /** Creates new FontFormatting */
        public PatternFormatting(RecordInputStream in1)
        {
            field_15_pattern_style = in1.ReadShort();
            field_16_pattern_color_indexes = in1.ReadShort();
        }

        // PATTERN FORMATING BLOCK
        // For Pattern Styles see constants at HSSFCellStyle (from NO_Fill to LEAST_DOTS)
        private short field_15_pattern_style;
        private static BitField FillPatternStyle = BitFieldFactory.GetInstance(0xFC00);

        private short field_16_pattern_color_indexes;
        private static BitField patternColorIndex = BitFieldFactory.GetInstance(0x007F);
        private static BitField patternBackgroundColorIndex = BitFieldFactory.GetInstance(0x3F80);

        /**
         * Get the Fill pattern 
         * @return Fill pattern
         */

        public short FillPattern
        {
            get
            {
                return FillPatternStyle.GetShortValue(field_15_pattern_style);
            }
            set 
            {
                field_15_pattern_style
                = FillPatternStyle.SetShortValue(field_15_pattern_style, value); 
            }
        }


        /**
         * Get the background Fill color
         * @see org.apache.poi.hssf.usermodel.HSSFPalette#GetColor(short)
         * @return Fill color
         */
        public short FillBackgroundColor
        {
            get
            {
                return patternBackgroundColorIndex.GetShortValue(field_16_pattern_color_indexes);
            }
            set { 
                field_16_pattern_color_indexes = 
                    patternBackgroundColorIndex.SetShortValue(field_16_pattern_color_indexes, value); 
            }
        }

        /**
         * Get the foreground Fill color
         * @see org.apache.poi.hssf.usermodel.HSSFPalette#GetColor(short)
         * @return Fill color
         */
        public short FillForegroundColor
        {
            get
            {
                return patternColorIndex.GetShortValue(field_16_pattern_color_indexes);
            }
            set 
            {
                field_16_pattern_color_indexes = patternColorIndex.SetShortValue(field_16_pattern_color_indexes, value);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("    [Pattern Formatting]\n");
            buffer.Append("          .Fillpattern= ").Append(StringUtil.ToHexString(FillPattern)).Append("\n");
            buffer.Append("          .fgcoloridx= ").Append(StringUtil.ToHexString(FillForegroundColor)).Append("\n");
            buffer.Append("          .bgcoloridx= ").Append(StringUtil.ToHexString(FillBackgroundColor)).Append("\n");
            buffer.Append("    [/Pattern Formatting]\n");
            return buffer.ToString();
        }

        public Object Clone()
        {
            PatternFormatting rec = new PatternFormatting();
            rec.field_15_pattern_style = field_15_pattern_style;
            rec.field_16_pattern_color_indexes = field_16_pattern_color_indexes;
            return rec;
        }

        public int Serialize(int offset, byte[] data)
        {
            LittleEndian.PutShort(data, offset, field_15_pattern_style);
            offset += 2;
            LittleEndian.PutShort(data, offset, field_16_pattern_color_indexes);
            offset += 2;
            return 4;
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_15_pattern_style);
            out1.WriteShort(field_16_pattern_color_indexes);
        }
    }
}
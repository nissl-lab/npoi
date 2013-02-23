
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

namespace NPOI.HSSF.Record.CF
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.SS.UserModel;

    /**
     * Font Formatting Block of the Conditional Formatting Rule Record.
     * 
     * @author Dmitriy Kumshayev
     */
    public class FontFormatting
    {
        private byte[] _rawData;

        private const int OFFSET_FONT_NAME = 0;
        private const int OFFSET_FONT_HEIGHT = 64;
        private const int OFFSET_FONT_OPTIONS = 68;
        private const int OFFSET_FONT_WEIGHT = 72;
        private const int OFFSET_ESCAPEMENT_TYPE = 74;
        private const int OFFSET_UNDERLINE_TYPE = 76;
        private const int OFFSET_FONT_COLOR_INDEX = 80;
        private const int OFFSET_OPTION_FLAGS = 88;
        private const int OFFSET_ESCAPEMENT_TYPE_MODIFIED = 92;
        private const int OFFSET_UNDERLINE_TYPE_MODIFIED = 96;
        private const int OFFSET_FONT_WEIGHT_MODIFIED = 100;
        private const int OFFSET_NOT_USED1 = 104;
        private const int OFFSET_NOT_USED2 = 108;
        private const int OFFSET_NOT_USED3 = 112; // for some reason Excel always Writes  0x7FFFFFFF at this offset   
        private const int OFFSET_FONT_FORMATING_END = 116;
        private const int RAW_DATA_SIZE = 118;


        public const int FONT_CELL_HEIGHT_PRESERVED = unchecked((int)0xFFFFFFFF);

        // FONT OPTIONS MASKS
        private static BitField posture = BitFieldFactory.GetInstance(0x00000002);
        private static BitField outline = BitFieldFactory.GetInstance(0x00000008);
        private static BitField shadow = BitFieldFactory.GetInstance(0x00000010);
        private static BitField cancellation = BitFieldFactory.GetInstance(0x00000080);

        // OPTION FLAGS MASKS

        private static BitField styleModified = BitFieldFactory.GetInstance(0x00000002);
        private static BitField outlineModified = BitFieldFactory.GetInstance(0x00000008);
        private static BitField shadowModified = BitFieldFactory.GetInstance(0x00000010);
        private static BitField cancellationModified = BitFieldFactory.GetInstance(0x00000080);

        /** Normal boldness (not bold) */
        private const short FONT_WEIGHT_NORMAL = 0x190;

        /**
         * Bold boldness (bold)
         */
        private const short FONT_WEIGHT_BOLD = 0x2bc;

        private FontFormatting(byte[] rawData)
        {
            _rawData = rawData;
        }

        public FontFormatting():this(new byte[RAW_DATA_SIZE])
        {
            

            FontHeight=-1;
            IsItalic=false;
            IsFontWeightModified=false;
            IsOutlineOn=false;
            IsShadowOn=false;
            IsStruckout=false;
            EscapementType=(FontSuperScript)0;
            UnderlineType=(FontUnderlineType)0;
            FontColorIndex=(short)-1;

            IsFontStyleModified=false;
            IsFontOutlineModified=false;
            IsFontShadowModified=false;
            IsFontCancellationModified=false;

            IsEscapementTypeModified=false;
            IsUnderlineTypeModified=false;

            SetShort(OFFSET_FONT_NAME, 0);
            SetInt(OFFSET_NOT_USED1, 0x00000001);
            SetInt(OFFSET_NOT_USED2, 0x00000000);
            SetInt(OFFSET_NOT_USED3, 0x7FFFFFFF);// for some reason Excel always Writes  0x7FFFFFFF at this offset
            SetShort(OFFSET_FONT_FORMATING_END, 0x0001);
        }

        /** Creates new FontFormatting */
        public FontFormatting(RecordInputStream in1):this(new byte[RAW_DATA_SIZE])
        {
            
            for (int i = 0; i < _rawData.Length; i++)
            {
                _rawData[i] =(byte) in1.ReadByte();
            }
        }

        private short GetShort(int offset)
        {
            return LittleEndian.GetShort(_rawData, offset);
        }
        private void SetShort(int offset, int value)
        {
            LittleEndian.PutShort(_rawData, offset, (short)value);
        }
        private int GetInt(int offset)
        {
            return LittleEndian.GetInt(_rawData, offset);
        }
        private void SetInt(int offset, int value)
        {
            LittleEndian.PutInt(_rawData, offset, value);
        }

        public byte[] GetRawRecord()
        {
            return _rawData;
        }

        /**
         * Gets the height of the font in 1/20th point Units
         *
         * @return fontheight (in points/20); or -1 if not modified
         */
        public int FontHeight
        {
            get{return GetInt(OFFSET_FONT_HEIGHT);}
            set { SetInt(OFFSET_FONT_HEIGHT, value); }
        }

        private void SetFontOption(bool option, BitField field)
        {
            int options = GetInt(OFFSET_FONT_OPTIONS);
            options = field.SetBoolean(options, option);
            SetInt(OFFSET_FONT_OPTIONS, options);
        }

        private bool GetFontOption(BitField field)
        {
            int options = GetInt(OFFSET_FONT_OPTIONS);
            return field.IsSet(options);
        }


        /**
         * Get whether the font Is to be italics or not
         *
         * @return italics - whether the font Is italics or not
         * @see #GetAttributes()
         */

        public bool IsItalic
        {
            get
            {
                return GetFontOption(posture);
            }
            set
            {
                SetFontOption(value, posture);
            }
        }

        public bool IsOutlineOn
        {
            get
            {
                return GetFontOption(outline);
            }
            set { SetFontOption(value, outline); }
        }

        public bool IsShadowOn
        {
            get
            {
                return GetFontOption(shadow);
            }
            set { SetFontOption(value, shadow); }
        }


        /**
         * Get whether the font Is to be stricken out or not
         *
         * @return strike - whether the font Is stricken out or not
         * @see #GetAttributes()
         */

        public bool IsStruckout
        {
            get
            {
                return GetFontOption(cancellation);
            }
            set { SetFontOption(value, cancellation); }
        }


        /// <summary>
        /// Get or set the font weight for this font (100-1000dec or 0x64-0x3e8).  
        /// Default Is 0x190 for normal and 0x2bc for bold
        /// </summary>
        public short FontWeight
        {
            get { return GetShort(OFFSET_FONT_WEIGHT); }
            set
            {
                short bw = value;
                if (bw < 100) { bw = 100; }
                if (bw > 1000) { bw = 1000; }
                SetShort(OFFSET_FONT_WEIGHT, bw);
            }
        }


        /// <summary>
        ///Get or set whether the font weight is set to bold or not 
        /// </summary>
        public bool IsBold
        {
            get
            {
                return FontWeight == FONT_WEIGHT_BOLD;
            }
            set { this.FontWeight = (value ? FONT_WEIGHT_BOLD : FONT_WEIGHT_NORMAL); }
        }

        /**
         * Get the type of base or subscript for the font
         *
         * @return base or subscript option
         * @see org.apache.poi.hssf.usermodel.HSSFFontFormatting#SS_NONE
         * @see org.apache.poi.hssf.usermodel.HSSFFontFormatting#SS_SUPER
         * @see org.apache.poi.hssf.usermodel.HSSFFontFormatting#SS_SUB
         */
        public FontSuperScript EscapementType
        {
            get
            {
                return (FontSuperScript)GetShort(OFFSET_ESCAPEMENT_TYPE);
            }
            set { SetShort(OFFSET_ESCAPEMENT_TYPE, (short)value); }
        }

        /**
         * Get the type of Underlining for the font
         *
         * @return font Underlining type
         */

        public FontUnderlineType UnderlineType
        {
            get
            {
                return (FontUnderlineType)GetShort(OFFSET_UNDERLINE_TYPE);
            }
            set { SetShort(OFFSET_UNDERLINE_TYPE, (short)value); }
        }



        public short FontColorIndex
        {
            get
            {
                return (short)GetInt(OFFSET_FONT_COLOR_INDEX);
            }
            set { SetInt(OFFSET_FONT_COLOR_INDEX, value); }
        }


        private bool GetOptionFlag(BitField field)
        {
            int optionFlags = GetInt(OFFSET_OPTION_FLAGS);
            int value = field.GetValue(optionFlags);
            return value == 0 ? true : false;
        }

        private void SetOptionFlag(bool modified, BitField field)
        {
            int value = modified ? 0 : 1;
            int optionFlags = GetInt(OFFSET_OPTION_FLAGS);
            optionFlags = field.SetValue(optionFlags, value);
            SetInt(OFFSET_OPTION_FLAGS, optionFlags);
        }


        public bool IsFontStyleModified
        {
            get { return GetOptionFlag(styleModified); }
            set { SetOptionFlag(value, styleModified); }
        }

        public bool IsFontOutlineModified
        {
            get { return GetOptionFlag(outlineModified); }
            set { SetOptionFlag(value, outlineModified); }
        }

        public bool IsFontShadowModified
        {
            get { return GetOptionFlag(shadowModified); }
            set { SetOptionFlag(value, shadowModified); }
        }

        public bool IsFontCancellationModified
        {
            get { return GetOptionFlag(cancellationModified); }
            set { SetOptionFlag(value, cancellationModified); }
        }

        public bool IsEscapementTypeModified
        {
            get
            {
                int escapementModified = GetInt(OFFSET_ESCAPEMENT_TYPE_MODIFIED);
                return escapementModified == 0;
            }
            set
            {
                int value1 = value ? 0 : 1;
                SetInt(OFFSET_ESCAPEMENT_TYPE_MODIFIED, value1);
            }
        }

        public bool IsUnderlineTypeModified
        {
            get
            {
                int underlineModified = GetInt(OFFSET_UNDERLINE_TYPE_MODIFIED);
                return underlineModified == 0;
            }
            set {

                int value1 = value ? 0 : 1;
                SetInt(OFFSET_UNDERLINE_TYPE_MODIFIED, value1);
            }
        }

        public bool IsFontWeightModified
        {
            get
            {
                int fontStyleModified = GetInt(OFFSET_FONT_WEIGHT_MODIFIED);
                return fontStyleModified == 0;
            }
            set
            {
                int value1 = value ? 0 : 1;
                SetInt(OFFSET_FONT_WEIGHT_MODIFIED, value1);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("	[Font Formatting]\n");

            buffer.Append("	.font height = ").Append(FontHeight).Append(" twips\n");

            if (IsFontStyleModified)
            {
                buffer.Append("	.font posture = ").Append(IsItalic ? "Italic" : "Normal").Append("\n");
            }
            else
            {
                buffer.Append("	.font posture = ]not modified]").Append("\n");
            }

            if (IsFontOutlineModified)
            {
                buffer.Append("	.font outline = ").Append(IsOutlineOn).Append("\n");
            }
            else
            {
                buffer.Append("	.font outline Is not modified\n");
            }

            if (IsFontShadowModified)
            {
                buffer.Append("	.font shadow = ").Append(IsShadowOn).Append("\n");
            }
            else
            {
                buffer.Append("	.font shadow Is not modified\n");
            }

            if (IsFontCancellationModified)
            {
                buffer.Append("	.font strikeout = ").Append(IsStruckout).Append("\n");
            }
            else
            {
                buffer.Append("	.font strikeout Is not modified\n");
            }

            if (IsFontStyleModified)
            {
                buffer.Append("	.font weight = ").
                    Append(FontWeight).
                    Append(
                        FontWeight == FONT_WEIGHT_NORMAL ? "(Normal)"
                                : FontWeight == FONT_WEIGHT_BOLD ? "(Bold)" : "0x" + StringUtil.ToHexString(FontWeight)).
                    Append("\n");
            }
            else
            {
                buffer.Append("	.font weight = ]not modified]").Append("\n");
            }

            if (IsEscapementTypeModified)
            {
                buffer.Append("	.escapement type = ").Append(EscapementType).Append("\n");
            }
            else
            {
                buffer.Append("	.escapement type Is not modified\n");
            }

            if (IsUnderlineTypeModified)
            {
                buffer.Append("	.underline type = ").Append(UnderlineType).Append("\n");
            }
            else
            {
                buffer.Append("	.underline type Is not modified\n");
            }
            buffer.Append("	.color index = ").Append("0x" + StringUtil.ToHexString(FontColorIndex).ToUpper()).Append("\n");

            buffer.Append("	[/Font Formatting]\n");
            return buffer.ToString();
        }

        public Object Clone()
        {
            byte[] rawData = (byte[])_rawData.Clone();
            return new FontFormatting(rawData);
        }
    }
}
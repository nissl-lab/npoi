
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
    using NPOI.SS.UserModel;
    using NPOI.Util;



    /**
     * Title:        Font Record - descrbes a font in the workbook (index = 0-3,5-infinity - skip 4)
     * Description:  An element in the Font Table
     * REFERENCE:  PG 315 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class FontRecord
       : StandardRecord
    {
        public const short sid = 0x31;            // docs are wrong (0x231 Microsoft Support site article Q184647)

        private short field_1_font_height;        // in Units of .05 of a point
        private short field_2_attributes;

        // 0 0x01 - Reserved bit must be 0
        static private BitField italic = BitFieldFactory.GetInstance(0x02);                // Is this font in italics

        // 2 0x04 - reserved bit must be 0
        static private BitField strikeout = BitFieldFactory.GetInstance(0x08);    //is this font has a line through the center
        static private BitField macoutline = BitFieldFactory.GetInstance(0x10);   // some weird macintosh thing....but who Understands those mac people anyhow
        static private BitField macshadow = BitFieldFactory.GetInstance(0x20);      // some weird macintosh thing....but who Understands those mac people anyhow

        // 7-6 - reserved bits must be 0
        // the rest Is Unused
        private short field_3_color_palette_index;
        private short field_4_bold_weight;
        private short field_5_base_sub_script;   // 00none/01base/02sub
        private byte field_6_underline;          // 00none/01single/02double/21singleaccounting/22doubleaccounting
        private byte field_7_family;             // ?? defined by windows api logfont structure?
        private byte field_8_charset;            // ?? defined by windows api logfont structure?
        private byte field_9_zero = 0;           // must be 0
        private String field_11_font_name;         // whoa...the font name

        public FontRecord()
        {
        }

        /**
         * Constructs a Font record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public FontRecord(RecordInputStream in1)
        {
            field_1_font_height = in1.ReadShort();
            field_2_attributes = in1.ReadShort();
            field_3_color_palette_index = in1.ReadShort();
            field_4_bold_weight = in1.ReadShort();
            field_5_base_sub_script = in1.ReadShort();
            field_6_underline = (byte)in1.ReadByte();
            field_7_family = (byte)in1.ReadByte();
            field_8_charset = (byte)in1.ReadByte();
            field_9_zero = (byte)in1.ReadByte();
            int field_10_font_name_len = (byte)in1.ReadByte();
            int unicodeFlags = in1.ReadUByte(); // options byte present always (even if no character data)

            if (field_10_font_name_len > 0)
            {
                if (unicodeFlags == 0)
                {   // Is compressed Unicode
                    field_11_font_name = in1.ReadCompressedUnicode(field_10_font_name_len);
                }
                else
                {   // Is not compressed Unicode
                    field_11_font_name = in1.ReadUnicodeLEString(field_10_font_name_len);
                }
            }
            else
            {
                field_11_font_name = "";
            }
        }

        /**
         * Clones all the font style information from another
         *  FontRecord, onto this one. This 
         *  will then hold all the same font style options.
         */
        public void CloneStyleFrom(FontRecord source)
        {
            field_1_font_height = source.field_1_font_height;
            field_2_attributes = source.field_2_attributes;
            field_3_color_palette_index = source.field_3_color_palette_index;
            field_4_bold_weight = source.field_4_bold_weight;
            field_5_base_sub_script = source.field_5_base_sub_script;
            field_6_underline = source.field_6_underline;
            field_7_family = source.field_7_family;
            field_8_charset = source.field_8_charset;
            field_9_zero = source.field_9_zero;
            field_11_font_name = source.field_11_font_name;
        }
        // attributes bitfields

        /**
         * Set the font to be italics or not
         *
         * @param italics - whether the font Is italics or not
         * @see #SetAttributes(short)
         */

        public bool IsItalic
        {
            set
            {
                field_2_attributes = italic.SetShortBoolean(field_2_attributes, value);
            }
            get
            {
                return italic.IsSet(field_2_attributes);
            }
        }

        /**
         * Set the font to be stricken out or not
         *
         * @param strike - whether the font Is stricken out or not
         * @see #SetAttributes(short)
         */

        public bool IsStrikeout
        {
            set { field_2_attributes = strikeout.SetShortBoolean(field_2_attributes, value); }
            get { return strikeout.IsSet(field_2_attributes); }
        }

        /**
         * whether to use the mac outline font style thing (mac only) - Some mac person
         * should comment this instead of me doing it (since I have no idea)
         *
         * @param mac - whether to do that mac font outline thing or not
         * @see #SetAttributes(short)
         */

        public bool IsMacoutlined
        {
            set { field_2_attributes = macoutline.SetShortBoolean(field_2_attributes, value); }
            get { return macoutline.IsSet(field_2_attributes); }
        }

        /**
         * whether to use the mac shado font style thing (mac only) - Some mac person
         * should comment this instead of me doing it (since I have no idea)
         *
         * @param mac - whether to do that mac font shadow thing or not
         * @see #SetAttributes(short)
         */

        public bool IsMacshadowed
        {
            set { field_2_attributes = macshadow.SetShortBoolean(field_2_attributes, value); }
            get { return macshadow.IsSet(field_2_attributes); }
        }

        /**
         * Set the type of Underlining for the font
         */

        public FontUnderlineType Underline
        {
            get { return (FontUnderlineType) field_6_underline; }
            set { field_6_underline = (byte) value; }
        }

        /**
         * Set the font family (TODO)
         *
         * @param f family
         */

        public byte Family
        {
            set { field_7_family = value; }
            get { return field_7_family; }
        }

        /**
         * Set the Char Set
         *
         * @param charSet - CharSet
         */

        public byte Charset
        {
            set { field_8_charset = value; }
            get { return field_8_charset; }
        }

        /**
         * Set the name of the font
         *
         * @param fn - name of the font (i.e. "Arial")
         */

        public string FontName
        {
            set { field_11_font_name = value; }
            get { return field_11_font_name; }
        }

        /**
         * Gets the height of the font in 1/20th point Units
         *
         * @return fontheight (in points/20)
         */

        public short FontHeight
        {
            get { return field_1_font_height; }
            set { field_1_font_height = value; }
        }

        /**
         * Get the font attributes (see individual bit Getters that reference this method)
         *
         * @return attribute - the bitmask
         */

        public short Attributes
        {
            get { return field_2_attributes; }
            set { field_2_attributes = value; }
        }

        /**
         * Get the font's color palette index
         *
         * @return cpi - font color index
         */

        public short ColorPaletteIndex
        {
            get { return field_3_color_palette_index; }
            set { field_3_color_palette_index = value; }
        }

        /**
         * Get the bold weight for this font (100-1000dec or 0x64-0x3e8).  Default Is
         * 0x190 for normal and 0x2bc for bold
         *
         * @return bw - a number between 100-1000 for the fonts "boldness"
         */

        public short BoldWeight
        {
            get { return field_4_bold_weight; }
            set { field_4_bold_weight = value; }
        }

        /**
         * Get the type of base or subscript for the font
         *
         * @return base or subscript option
         */

        public FontSuperScript SuperSubScript
        {
            get { return (FontSuperScript) field_5_base_sub_script; }
            set { field_5_base_sub_script = (short) value; }
        }
        /**
 * Does this FontRecord have all the same font
 *  properties as the supplied FontRecord?
 * Note that {@link #equals(Object)} will check
 *  for exact objects, while this will check
 *  for exact contents, because normally the
 *  font record's position makes a big
 *  difference too.  
 */
        public bool SameProperties(FontRecord other)
        {
            return
            field_1_font_height == other.field_1_font_height &&
            field_2_attributes == other.field_2_attributes &&
            field_3_color_palette_index == other.field_3_color_palette_index &&
            field_4_bold_weight == other.field_4_bold_weight &&
            field_5_base_sub_script == other.field_5_base_sub_script &&
            field_6_underline == other.field_6_underline &&
            field_7_family == other.field_7_family &&
            field_8_charset == other.field_8_charset &&
            field_9_zero == other.field_9_zero &&
            field_11_font_name.Equals(other.field_11_font_name);
        }
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FONT]\n");
            buffer.Append("    .fontheight      = ")
                .Append(StringUtil.ToHexString(FontHeight)).Append("\n");
            buffer.Append("    .attributes      = ")
                .Append(StringUtil.ToHexString(Attributes)).Append("\n");
            buffer.Append("         .italic     = ").Append(IsItalic)
                .Append("\n");
            buffer.Append("         .strikout   = ").Append(IsStrikeout)
                .Append("\n");
            buffer.Append("         .macoutlined= ").Append(IsMacoutlined)
                .Append("\n");
            buffer.Append("         .macshadowed= ").Append(IsMacshadowed)
                .Append("\n");
            buffer.Append("    .colorpalette    = ")
                .Append(StringUtil.ToHexString(ColorPaletteIndex)).Append("\n");
            buffer.Append("    .boldweight      = ")
                .Append(StringUtil.ToHexString(BoldWeight)).Append("\n");
            buffer.Append("    .basesubscript  = ")
                .Append(StringUtil.ToHexString((short) SuperSubScript)).Append("\n");
            buffer.Append("    .underline       = ")
                .Append(StringUtil.ToHexString((short) Underline)).Append("\n");
            buffer.Append("    .family          = ")
                .Append(StringUtil.ToHexString(Family)).Append("\n");
            buffer.Append("    .charset         = ")
                .Append(StringUtil.ToHexString(Charset)).Append("\n");
            buffer.Append("    .fontname        = ").Append(FontName)
                .Append("\n");
            buffer.Append("[/FONT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(FontHeight);
            out1.WriteShort(Attributes);
            out1.WriteShort(ColorPaletteIndex);
            out1.WriteShort(BoldWeight);
            out1.WriteShort((int) SuperSubScript);
            out1.WriteByte((int) Underline);
            out1.WriteByte(Family);
            out1.WriteByte(Charset);
            out1.WriteByte(field_9_zero);
            int fontNameLen = field_11_font_name.Length;
            out1.WriteByte(fontNameLen);
            bool hasMultibyte = StringUtil.HasMultibyte(field_11_font_name);
            out1.WriteByte(hasMultibyte ? 0x01 : 0x00);
            if (fontNameLen > 0)
            {
                if (hasMultibyte)
                {
                    StringUtil.PutUnicodeLE(field_11_font_name, out1);
                }
                else
                {
                    StringUtil.PutCompressedUnicode(field_11_font_name, out1);
                }
            }

        }

        protected override int DataSize
        {
            get
            {
                int size = 16; // 5 shorts + 6 bytes
                int fontNameLen = field_11_font_name.Length;
                if (fontNameLen < 1)
                {
                    return size;
                }

                bool hasMultibyte = StringUtil.HasMultibyte(field_11_font_name);
                return size + fontNameLen * (hasMultibyte ? 2 : 1);
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override int GetHashCode()
        {
            int prime = 31;
            int result = 1;
            result = prime
                    * result
                    + ((field_11_font_name == null) ? 0 : field_11_font_name
                            .GetHashCode());
            result = prime * result + field_1_font_height;
            result = prime * result + field_2_attributes;
            result = prime * result + field_3_color_palette_index;
            result = prime * result + field_4_bold_weight;
            result = prime * result + field_5_base_sub_script;
            result = prime * result + field_6_underline;
            result = prime * result + field_7_family;
            result = prime * result + field_8_charset;
            result = prime * result + field_9_zero;
            return result;
        }

        /**
         * Only returns two for the same exact object -
         *  creating a second FontRecord with the same
         *  properties won't be considered equal, as 
         *  the record's position in the record stream
         *  matters.
         */
        public override bool Equals(Object obj)
        {
            if (this == obj)
                return true;
            return false;
        }
    }
}

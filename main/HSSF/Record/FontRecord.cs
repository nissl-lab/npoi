
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
    using System.IO;
    using System.Text;
    using System.Collections;
    using NPOI.Util;


    /**
     * Title:        Font Record - descrbes a font in the workbook (index = 0-3,5-infinity - skip 4)
     * Description:  An element in the Font Table
     * REFERENCE:  PG 315 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class FontRecord
       : Record
    {
        public const short sid =
            0x31;                                                 // docs are wrong (0x231 Microsoft Support site article Q184647)
        public static short SS_NONE = 0;
        public static short SS_SUPER = 1;
        public static short SS_SUB = 2;
        public static byte U_NONE = 0;
        public static byte U_SINGLE = 1;
        public static byte U_DOUBLE = 2;
        public static byte U_SINGLE_ACCOUNTING = 0x21;
        public static byte U_DOUBLE_ACCOUNTING = 0x22;
        private short field_1_font_height;        // in Units of .05 of a point
        private short field_2_attributes;

        // 0 0x01 - Reserved bit must be 0
        static private BitField italic =
            BitFieldFactory.GetInstance(0x02);                                   // Is this font in italics

        // 2 0x04 - reserved bit must be 0
        static private BitField strikeout =
            BitFieldFactory.GetInstance(0x08);                                   // Is this font has a line through the center
        static private BitField macoutline = BitFieldFactory.GetInstance(
            0x10);                                                // some weird macintosh thing....but who Understands those mac people anyhow
        static private BitField macshadow = BitFieldFactory.GetInstance(
            0x20);                                                // some weird macintosh thing....but who Understands those mac people anyhow

        // 7-6 - reserved bits must be 0
        // the rest Is Unused
        private short field_3_color_palette_index;
        private short field_4_bold_weight;
        private short field_5_base_sub_script;   // 00none/01base/02sub
        private byte field_6_underline;          // 00none/01single/02double/21singleaccounting/22doubleaccounting
        private byte field_7_family;             // ?? defined by windows api logfont structure?
        private byte field_8_charset;            // ?? defined by windows api logfont structure?
        private byte field_9_zero = 0;           // must be 0
        private byte field_10_font_name_len;     // Length of the font name
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
            field_10_font_name_len = (byte)in1.ReadByte();
            if (field_10_font_name_len > 0)
            {
                if (in1.ReadByte() == 0)
                {   // Is compressed Unicode
                    field_11_font_name = in1.ReadCompressedUnicode(LittleEndian.UByteToInt(field_10_font_name_len));
                }
                else
                {   // Is not compressed Unicode
                    field_11_font_name = in1.ReadUnicodeLEString(field_10_font_name_len);
                }
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
            field_10_font_name_len = source.field_10_font_name_len;
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
            set{field_2_attributes = strikeout.SetShortBoolean(field_2_attributes, value);}
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
         *
         * @param u  base or subscript option
         *
         * @see #U_NONE
         * @see #U_SINGLE
         * @see #U_DOUBLE
         * @see #U_SINGLE_ACCOUNTING
         * @see #U_DOUBLE_ACCOUNTING
         */

        public byte Underline
        {
            set { field_6_underline = value; }
            get { return field_6_underline; }
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

        public byte CharSet
        {
            set { field_8_charset = value; }
            get { return field_8_charset; }
        }

        /**
         * Set the Length of the fontname string
         *
         * @param len  Length of the font name
         * @see #SetFontName(String)
         */

        public byte FontNameLength
        {
            set { field_10_font_name_len = value; }
            get { return field_10_font_name_len; }
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
            get{return field_1_font_height;}
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
            get{return field_3_color_palette_index;}
            set{field_3_color_palette_index =value;}
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
         * @see #SS_NONE
         * @see #SS_SUPER
         * @see #SS_SUB
         */

        public short SuperSubScript
        {
            get { return field_5_base_sub_script; }
            set { field_5_base_sub_script = value; }
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
            field_10_font_name_len == other.field_10_font_name_len &&
            field_11_font_name.Equals(other.field_11_font_name)
            ;
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
                .Append(StringUtil.ToHexString(SuperSubScript)).Append("\n");
            buffer.Append("    .underline       = ")
                .Append(StringUtil.ToHexString(Underline)).Append("\n");
            buffer.Append("    .family          = ")
                .Append(StringUtil.ToHexString(Family)).Append("\n");
            buffer.Append("    .charSet         = ")
                .Append(StringUtil.ToHexString(CharSet)).Append("\n");
            buffer.Append("    .nameLength      = ")
                .Append(StringUtil.ToHexString(FontNameLength)).Append("\n");
            buffer.Append("    .fontname        = ").Append(FontName)
                .Append("\n");
            buffer.Append("[/FONT]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            int realflen = FontNameLength * 2;

            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(
                data, 2 + offset,
                (short)(15 + realflen
                           + 1));   // 19 - 4 (sid/len) + font name Length = datasize

            // Undocumented single byte (1)
            LittleEndian.PutShort(data, 4 + offset, FontHeight);
            LittleEndian.PutShort(data, 6 + offset, Attributes);
            LittleEndian.PutShort(data, 8 + offset, ColorPaletteIndex);
            LittleEndian.PutShort(data, 10 + offset, BoldWeight);
            LittleEndian.PutShort(data, 12 + offset, SuperSubScript);
            data[14 + offset] = Underline;
            data[15 + offset] = Family;
            data[16 + offset] = CharSet;
            data[17 + offset] = field_9_zero;
            data[18 + offset] = FontNameLength;
            data[19 + offset] = (byte)1;
            if (FontName != null)
            {
                StringUtil.PutUnicodeLE(FontName, data, 20 + offset);
            }
            return RecordSize;
        }

        public override int RecordSize
        {
            get { return (FontNameLength * 2) + 20; }
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
            result = prime * result + field_10_font_name_len;
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
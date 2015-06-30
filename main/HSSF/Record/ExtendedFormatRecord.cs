
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
     * Title:        Extended Format Record
     * Description:  Probably one of the more complex records.  There are two breeds:
     *               Style and Cell.
     *
     *               It should be noted that fields in the extended format record are
     *               somewhat arbitrary.  Almost all of the fields are bit-level, but
     *               we name them as best as possible by functional Group.  In some
     *               places this Is better than others.
     *
     *
     * REFERENCE:  PG 426 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class ExtendedFormatRecord : StandardRecord
    {
        public const short sid = 0xE0;

        // null constant
        public const short NULL = unchecked((short)0xfff0);

        // xf type
        public const short XF_STYLE = 1;
        public const short XF_CELL = 0;

        // borders
        public const short NONE = 0x0;
        public const short THIN = 0x1;
        public const short MEDIUM = 0x2;
        public const short DASHED = 0x3;
        public const short DOTTED = 0x4;
        public const short THICK = 0x5;
        public const short DOUBLE = 0x6;
        public const short HAIR = 0x7;
        public const short MEDIUM_DASHED = 0x8;
        public const short DASH_DOT = 0x9;
        public const short MEDIUM_DASH_DOT = 0xA;
        public const short DASH_DOT_DOT = 0xB;
        public const short MEDIUM_DASH_DOT_DOT = 0xC;
        public const short SLANTED_DASH_DOT = 0xD;

        // alignment
        public const short GENERAL = 0x0;
        public const short LEFT = 0x1;
        public const short CENTER = 0x2;
        public const short RIGHT = 0x3;
        public const short FILL = 0x4;
        public const short JUSTIFY = 0x5;
        public const short CENTER_SELECTION = 0x6;

        // vertical alignment
        public const short VERTICAL_TOP = 0x0;
        public const short VERTICAL_CENTER = 0x1;
        public const short VERTICAL_BOTTOM = 0x2;
        public const short VERTICAL_JUSTIFY = 0x3;

        // fill
        public const short NO_FILL = 0;
        public const short SOLID_FILL = 1;
        public const short FINE_DOTS = 2;
        public const short ALT_BARS = 3;
        public const short SPARSE_DOTS = 4;
        public const short THICK_HORZ_BANDS = 5;
        public const short THICK_VERT_BANDS = 6;
        public const short THICK_BACKWARD_DIAG = 7;
        public const short THICK_FORWARD_DIAG = 8;
        public const short BIG_SPOTS = 9;
        public const short BRICKS = 10;
        public const short THIN_HORZ_BANDS = 11;
        public const short THIN_VERT_BANDS = 12;
        public const short THIN_BACKWARD_DIAG = 13;
        public const short THIN_FORWARD_DIAG = 14;
        public const short SQUARES = 15;
        public const short DIAMONDS = 16;

        // fields in BOTH style and Cell XF records
        private short field_1_font_index;             // not bit-mapped
        private short field_2_format_index;           // not bit-mapped

        // field_3_cell_options bit map
        static private BitField _locked = BitFieldFactory.GetInstance(0x0001);
        static private BitField _hidden = BitFieldFactory.GetInstance(0x0002);
        static private BitField _xf_type = BitFieldFactory.GetInstance(0x0004);
        static private BitField _123_prefix = BitFieldFactory.GetInstance(0x0008);
        static private BitField _parent_index = BitFieldFactory.GetInstance(0xFFF0);
        private short field_3_cell_options;

        // field_4_alignment_options bit map
        static private BitField _alignment = BitFieldFactory.GetInstance(0x0007);
        static private BitField _wrap_text = BitFieldFactory.GetInstance(0x0008);
        static private BitField _vertical_alignment = BitFieldFactory.GetInstance(0x0070);
        static private BitField _justify_last = BitFieldFactory.GetInstance(0x0080);
        static private BitField _rotation = BitFieldFactory.GetInstance(0xFF00);
        private short field_4_alignment_options;

        // field_5_indention_options
        static private BitField _indent =
            BitFieldFactory.GetInstance(0x000F);
        static private BitField _shrink_to_fit =
            BitFieldFactory.GetInstance(0x0010);
        static private BitField _merge_cells =
            BitFieldFactory.GetInstance(0x0020);
        static private BitField _Reading_order =
            BitFieldFactory.GetInstance(0x00C0);

        // apparently bits 8 and 9 are Unused
        static private BitField _indent_not_parent_format =
            BitFieldFactory.GetInstance(0x0400);
        static private BitField _indent_not_parent_font =
            BitFieldFactory.GetInstance(0x0800);
        static private BitField _indent_not_parent_alignment =
            BitFieldFactory.GetInstance(0x1000);
        static private BitField _indent_not_parent_border =
            BitFieldFactory.GetInstance(0x2000);
        static private BitField _indent_not_parent_pattern =
            BitFieldFactory.GetInstance(0x4000);
        static private BitField _indent_not_parent_cell_options =
            BitFieldFactory.GetInstance(0x8000);
        private short field_5_indention_options;

        // field_6_border_options bit map
        static private BitField _border_left = BitFieldFactory.GetInstance(0x000F);
        static private BitField _border_right = BitFieldFactory.GetInstance(0x00F0);
        static private BitField _border_top = BitFieldFactory.GetInstance(0x0F00);
        static private BitField _border_bottom = BitFieldFactory.GetInstance(0xF000);
        private short field_6_border_options;

        // all three of the following attributes are palette options
        // field_7_palette_options bit map
        static private BitField _left_border_palette_idx =
            BitFieldFactory.GetInstance(0x007F);
        static private BitField _right_border_palette_idx =
            BitFieldFactory.GetInstance(0x3F80);
        static private BitField _diag =
            BitFieldFactory.GetInstance(0xC000);
        private short field_7_palette_options;

        // field_8_adtl_palette_options bit map
        static private BitField _top_border_palette_idx =
            BitFieldFactory.GetInstance(0x0000007F);
        static private BitField _bottom_border_palette_idx =
            BitFieldFactory.GetInstance(0x00003F80);
        //is this used for diagional border color?
        static private BitField _adtl_diag_border_palette_idx =
            BitFieldFactory.GetInstance(0x001fc000);
        static private BitField _adtl_diag_line_style =
            BitFieldFactory.GetInstance(0x01e00000);

        // apparently bit 25 Is Unused
        static private BitField _adtl_fill_pattern =
            BitFieldFactory.GetInstance(unchecked((int)0xfc000000));
        private int field_8_adtl_palette_options;   // Additional to avoid 2

        // field_9_fill_palette_options bit map
        static private BitField _fill_foreground = BitFieldFactory.GetInstance(0x007F);
        static private BitField _fill_background = BitFieldFactory.GetInstance(0x3f80);

        // apparently bits 15 and 14 are Unused
        private short field_9_fill_palette_options;

        /**
         * Constructor ExtendedFormatRecord
         *
         *
         */

        public ExtendedFormatRecord()
        {
        }

        /**
         * Constructs an ExtendedFormat record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public ExtendedFormatRecord(RecordInputStream in1)
        {
            field_1_font_index = in1.ReadShort();
            field_2_format_index = in1.ReadShort();
            field_3_cell_options = in1.ReadShort();
            field_4_alignment_options = in1.ReadShort();
            field_5_indention_options = in1.ReadShort();
            field_6_border_options = in1.ReadShort();
            field_7_palette_options = in1.ReadShort();
            field_8_adtl_palette_options = in1.ReadInt();
            field_9_fill_palette_options = in1.ReadShort();
        }

        /**
 * Clones all the style information from another
 *  ExtendedFormatRecord, onto this one. This 
 *  will then hold all the same style options.
 *  
 * If The source ExtendedFormatRecord comes from
 *  a different Workbook, you will need to sort
 *  out the font and format indicies yourself!
 */
        public void CloneStyleFrom(ExtendedFormatRecord source)
        {
            field_1_font_index = source.field_1_font_index;
            field_2_format_index = source.field_2_format_index;
            field_3_cell_options = source.field_3_cell_options;
            field_4_alignment_options = source.field_4_alignment_options;
            field_5_indention_options = source.field_5_indention_options;
            field_6_border_options = source.field_6_border_options;
            field_7_palette_options = source.field_7_palette_options;
            field_8_adtl_palette_options = source.field_8_adtl_palette_options;
            field_9_fill_palette_options = source.field_9_fill_palette_options;
        }


        /// <summary>
        /// Get the index to the FONT record (which font to use 0 based)
        /// </summary>
        public short FontIndex
        {
            get { return field_1_font_index; }
            set { field_1_font_index = value; }
        }
        /// <summary>
        /// Get the index to the Format record (which FORMAT to use 0-based)
        /// </summary>
        public short FormatIndex
        {
            get
            {
                return field_2_format_index;
            }
            set { field_2_format_index = value; }
        }

        /// <summary>
        /// Gets the options bitmask - you can also use corresponding option bit Getters
        /// (see other methods that reference this one)
        /// </summary>
        public short CellOptions
        {
            get
            {
                return field_3_cell_options;
            }
            set { field_3_cell_options = value; }
        }

        /// <summary>
        /// Get whether the cell Is locked or not
        /// </summary>
        public bool IsLocked
        {
            get
            {
                return _locked.IsSet(field_3_cell_options);
            }
            set
            {
                field_3_cell_options = _locked.SetShortBoolean(field_3_cell_options,
                    value);
            }
        }

        /// <summary>
        /// Get whether the cell Is hidden or not
        /// </summary>
        public bool IsHidden
        {
            get
            {
                return _hidden.IsSet(field_3_cell_options);
            }
            set
            {
                field_3_cell_options = _hidden.SetShortBoolean(field_3_cell_options,
                    value);
            }
        }

        /// <summary>
        /// Get whether the cell Is a cell or style XFRecord
        /// </summary>
        public short XFType
        {
            get
            {
                return _xf_type.GetShortValue(field_3_cell_options);
            }
            set
            {
                field_3_cell_options = _xf_type.SetShortValue(field_3_cell_options,
                    value);
            }
        }


        /// <summary>
        /// Get some old holdover from lotus 123.  Who cares, its all over for Lotus.
        /// RIP Lotus.
        /// </summary>
        public bool _123Prefix
        {
            get{
                return _123_prefix.IsSet(field_3_cell_options);
            }
            set
            {
                field_3_cell_options =
                    _123_prefix.SetShortBoolean(field_3_cell_options, value);
            }
        }
        /// <summary>
        /// for cell XF types this Is the parent style (usually 0/normal).  For
        /// style this should be NULL.
        /// </summary>
        public short ParentIndex
        {
            get
            {
                return _parent_index.GetShortValue(field_3_cell_options);
            }
            set
            {
                field_3_cell_options =
                    _parent_index.SetShortValue(field_3_cell_options, value);
            }
        }

        /// <summary>
        /// Get the alignment options bitmask.  See corresponding bitGetter methods
        /// that reference this one.
        /// </summary>
        public short AlignmentOptions
        {
            get
            {
                return field_4_alignment_options;
            }
            set { field_4_alignment_options = value; }
        }


        /// <summary>
        /// Get the horizontal alignment of the cell.
        /// </summary>
        public short Alignment
        {
            get
            {
                return _alignment.GetShortValue(field_4_alignment_options);
            }
            set
            {
                field_4_alignment_options =
                    _alignment.SetShortValue(field_4_alignment_options, value);
            }
        }

        /// <summary>
        /// Get whether to wrap the text in the cell
        /// </summary>
        public bool WrapText
        {
            get
            {
                return _wrap_text.IsSet(field_4_alignment_options);
            }
            set
            {
                field_4_alignment_options =
                    _wrap_text.SetShortBoolean(field_4_alignment_options, value);
            }
        }

        /// <summary>
        /// Get the vertical alignment of text in the cell
        /// </summary>
        public short VerticalAlignment
        {
            get
            {
                return _vertical_alignment.GetShortValue(field_4_alignment_options);
            }
            set
            {
                field_4_alignment_options =
                    _vertical_alignment.SetShortValue(field_4_alignment_options,
                                    value);
            }
        }


        /// <summary>
        /// Docs just say this Is for far east versions..  (I'm guessing it
        /// justifies for right-to-left Read languages)
        /// </summary>
        public short JustifyLast
        {
            get
            {// for far east languages supported only for format always 0 for US
                return _justify_last.GetShortValue(field_4_alignment_options);
            }
            set
            { // for far east languages supported only for format always 0 for US
                field_4_alignment_options =
                    _justify_last.SetShortValue(field_4_alignment_options, value);
            }
        }

        /// <summary>
        /// Get the degree of rotation.  (I've not actually seen this used anywhere)
        /// </summary>
        public short Rotation
        {
            get
            {
                return _rotation.GetShortValue(field_4_alignment_options);
            }
            set
            {
                field_4_alignment_options =
                    _rotation.SetShortValue(field_4_alignment_options, value);
            }
        }

        /// <summary>
        /// Get the indent options bitmask  (see corresponding bit Getters that reference
        /// this field)
        /// </summary>
        public short IndentionOptions
        {
            get
            {
                return field_5_indention_options;
            }
            set { field_5_indention_options = value; }
        }

        /// <summary>
        ///  Get indention (not sure of the Units, think its spaces)
        /// </summary>
        public short Indent
        {
            get
            {
                return _indent.GetShortValue(field_5_indention_options);
            }
            set
            {
                field_5_indention_options =
                    _indent.SetShortValue(field_5_indention_options, value);
            }
        }


        /// <summary>
        /// Get whether to shrink the text to fit
        /// </summary>
        public bool ShrinkToFit
        {
            get
            {
                return _shrink_to_fit.IsSet(field_5_indention_options);
            }
            set
            {
                field_5_indention_options =
                    _shrink_to_fit.SetShortBoolean(field_5_indention_options, value);
            }
        }

        /// <summary>
        /// Get whether to merge cells
        /// </summary>
        public bool MergeCells
        {
            get
            {
                return _merge_cells.IsSet(field_5_indention_options);
            }
            set
            {
                field_5_indention_options =
                    _merge_cells.SetShortBoolean(field_5_indention_options, value);
            }
        }

        /// <summary>
        ///  Get the Reading order for far east versions (0 - Context, 1 - Left to right,
        /// 2 - right to left) - We could use some help with support for the far east.
        /// </summary>
        public short ReadingOrder
        {
            get
            {// only for far east  always 0 in US
                return _Reading_order.GetShortValue(field_5_indention_options);
            }
            set
            { // only for far east  always 0 in US
                field_5_indention_options =
                    _Reading_order.SetShortValue(field_5_indention_options, value);
            }
        }


        /// <summary>
        /// Get whether or not to use the format in this XF instead of the parent XF.
        /// </summary>
        public bool IsIndentNotParentFormat
        {
            get
            {
                return _indent_not_parent_format.IsSet(field_5_indention_options);
            }
            set
            {
                field_5_indention_options =
                    _indent_not_parent_format
                    .SetShortBoolean(field_5_indention_options, value);
            }

        }

        /// <summary>
        /// Get whether or not to use the font in this XF instead of the parent XF.
        /// </summary>
        public bool IsIndentNotParentFont
        {
            get
            {
                return _indent_not_parent_font.IsSet(field_5_indention_options);
            }
            set
            {
                field_5_indention_options =
                    _indent_not_parent_font.SetShortBoolean(field_5_indention_options,
                                          value);
            }
        }

        /// <summary>
        /// Get whether or not to use the alignment in this XF instead of the parent XF.
        /// </summary>
        public bool IsIndentNotParentAlignment
        {
            get{return _indent_not_parent_alignment.IsSet(field_5_indention_options);}
            set
            {
                field_5_indention_options =
                    _indent_not_parent_alignment
                    .SetShortBoolean(field_5_indention_options, value);
            }
        }

        /// <summary>
        /// Get whether or not to use the border in this XF instead of the parent XF.
        /// </summary>
        public bool IsIndentNotParentBorder
        {
            get { return _indent_not_parent_border.IsSet(field_5_indention_options); }
            set
            {
                field_5_indention_options =
                    _indent_not_parent_border
                    .SetShortBoolean(field_5_indention_options, value);
            }
        }

        
        /// <summary>
        /// Get whether or not to use the pattern in this XF instead of the parent XF.
        /// (foregrount/background)
        /// </summary>
        public bool IsIndentNotParentPattern
        {
            get { return _indent_not_parent_pattern.IsSet(field_5_indention_options); }
            set
            {
                field_5_indention_options =
                    _indent_not_parent_pattern
                    .SetShortBoolean(field_5_indention_options, value);
            }
        }

        /// <summary>
        /// Get whether or not to use the locking/hidden in this XF instead of the parent XF.
        /// </summary>
        public bool IsIndentNotParentCellOptions
        {
            get
            {
                return _indent_not_parent_cell_options
                    .IsSet(field_5_indention_options);
            }
            set
            {
                field_5_indention_options =
                    _indent_not_parent_cell_options
                    .SetShortBoolean(field_5_indention_options, value);
            }
        }

        /// <summary>
        /// Get the border options bitmask (see the corresponding bit Getter methods
        /// that reference back to this one)
        /// </summary>
        public short BorderOptions
        {
            get { return field_6_border_options; }
            set { field_6_border_options = value; }
        }


        /// <summary>
        /// Get the borderline style for the left border
        /// </summary>
        public short BorderLeft
        {
           get{return _border_left.GetShortValue(field_6_border_options);}
            set
            {
                field_6_border_options =
                    _border_left.SetShortValue(field_6_border_options, value);
            }
        }


        /// <summary>
        /// Get the borderline style for the right border
        /// </summary>
        public short BorderRight
        {
            get{return _border_right.GetShortValue(field_6_border_options);}
            set
            {
                field_6_border_options =
                    _border_right.SetShortValue(field_6_border_options, value);
            }
        }


        /// <summary>
        /// Get the borderline style for the top border
        /// </summary>
        public short BorderTop
        {
            get{return _border_top.GetShortValue(field_6_border_options);}
            set {
                field_6_border_options =_border_top.SetShortValue(field_6_border_options, value); 

            }
        }

        /// <summary>
        /// Get the borderline style for the bottom border
        /// </summary>
        public short BorderBottom
        {
            get{return _border_bottom.GetShortValue(field_6_border_options);}
            set {
                field_6_border_options =_border_bottom.SetShortValue(field_6_border_options, value);
            }
        }
        /// <summary>
        /// Get the palette options bitmask (see the individual bit Getter methods that
        /// reference this one) 
        /// </summary>
        public short PaletteOptions
        {
            get{return field_7_palette_options;}
            set { field_7_palette_options = value; }
        }

        /// <summary>
        ///  Get the palette index for the left border color
        /// </summary>
        public short LeftBorderPaletteIdx
        {
            get{return _left_border_palette_idx
                .GetShortValue(field_7_palette_options);
            }
            set {
                field_7_palette_options =
        _left_border_palette_idx.SetShortValue(field_7_palette_options,
                                               value);
            }
        }

        
        /// <summary>
        /// Get the palette index for the right border color
        /// </summary>
        public short RightBorderPaletteIdx
        {
            get{return _right_border_palette_idx
                .GetShortValue(field_7_palette_options);
            }
            set
            {
                field_7_palette_options =
                    _right_border_palette_idx.SetShortValue(field_7_palette_options,
                                          value);
            }
        }


        /// <summary>
        /// Get the Additional palette options bitmask (see individual bit Getter methods
        /// that reference this method)
        /// </summary>
        public int AdtlPaletteOptions
        {
            get{return field_8_adtl_palette_options;}
            set { field_8_adtl_palette_options = value; }
        }

        /// <summary>
        /// Get the palette index for the top border
        /// </summary>
        public short TopBorderPaletteIdx
        {
            get{return (short)_top_border_palette_idx
                .GetValue(field_8_adtl_palette_options);}
            set
            {
                field_8_adtl_palette_options =
                    _top_border_palette_idx.SetValue(field_8_adtl_palette_options,
                                   value);
            }
        }

        /// <summary>
        /// Get the palette index for the bottom border
        /// </summary>
        public short BottomBorderPaletteIdx
        {
            get{return (short)_bottom_border_palette_idx
                .GetValue(field_8_adtl_palette_options);
            }
            set
            {
                field_8_adtl_palette_options =
                    _bottom_border_palette_idx.SetValue(field_8_adtl_palette_options,
                                      value);
            }
        }

        /// <summary>
        /// Get for diagonal borders
        /// </summary>
        public short AdtlDiagBorderPaletteIdx
        {
            get{return (short)_adtl_diag_border_palette_idx.GetValue(field_8_adtl_palette_options);}
            set
            {
                field_8_adtl_palette_options =
                    _adtl_diag_border_palette_idx.SetValue(field_8_adtl_palette_options, value);
            }
        }

         
        /// <summary>
        /// Get the diagonal border line style
        /// </summary>
        public short AdtlDiagLineStyle
        {
            get{return (short)_adtl_diag_line_style
                .GetValue(field_8_adtl_palette_options);}
            set
            {
                field_8_adtl_palette_options =
                    _adtl_diag_line_style.SetValue(field_8_adtl_palette_options,
                                 value);
            }
        }
        /// <summary>
        /// Not sure what this Is for (maybe Fill lines?) 1 = down, 2 = up, 3 = both, 0 for none..
        /// </summary>
        public short Diagonal
        {
            get { return _diag.GetShortValue(field_7_palette_options); }
            set
            {
                field_7_palette_options = _diag.SetShortValue(field_7_palette_options,
                    value);
            }
        }
        /// <summary>
        /// Get the Additional Fill pattern
        /// </summary>
        public short AdtlFillPattern
        {
            get{return (short)_adtl_fill_pattern
                .GetValue(field_8_adtl_palette_options);}
            set
            {
                field_8_adtl_palette_options =
                    _adtl_fill_pattern.SetValue(field_8_adtl_palette_options, value);
            }
        }

        /// <summary>
        /// Get the Fill palette options bitmask (see indivdual bit Getters that
        /// reference this method)
        /// </summary>
        public short FillPaletteOptions
        {
            get{return field_9_fill_palette_options;}
            set { field_9_fill_palette_options = value; }
        }

        /// <summary>
        /// Get the foreground palette color index
        /// </summary>
        public short FillForeground
        {
            get{return _fill_foreground.GetShortValue(field_9_fill_palette_options);}
            set
            {
                field_9_fill_palette_options =
                    _fill_foreground.SetShortValue(field_9_fill_palette_options,
                                 value);
            }
        }

        /// <summary>
        /// Get the background palette color index
        /// </summary>
        public short FillBackground
        {
            get{return _fill_background.GetShortValue(field_9_fill_palette_options);}
            set
            {
                field_9_fill_palette_options =
                    _fill_background.SetShortValue(field_9_fill_palette_options,
                                 value);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[EXTENDEDFORMAT]\n");
            if (XFType == XF_STYLE)
            {
                buffer.Append(" STYLE_RECORD_TYPE\n");
            }
            else if (XFType == XF_CELL)
            {
                buffer.Append(" CELL_RECORD_TYPE\n");
            }
            buffer.Append("    .fontindex       = ")
                .Append(StringUtil.ToHexString(FontIndex)).Append("\n");
            buffer.Append("    .formatindex     = ")
                .Append(StringUtil.ToHexString(FormatIndex)).Append("\n");
            buffer.Append("    .celloptions     = ")
                .Append(StringUtil.ToHexString(CellOptions)).Append("\n");
            buffer.Append("          .Islocked  = ").Append(IsLocked)
                .Append("\n");
            buffer.Append("          .Ishidden  = ").Append(IsHidden)
                .Append("\n");
            buffer.Append("          .recordtype= ")
                .Append(StringUtil.ToHexString(XFType)).Append("\n");
            buffer.Append("          .parentidx = ")
                .Append(StringUtil.ToHexString(ParentIndex)).Append("\n");
            buffer.Append("    .alignmentoptions= ")
                .Append(StringUtil.ToHexString(AlignmentOptions)).Append("\n");
            buffer.Append("          .alignment = ").Append(Alignment)
                .Append("\n");
            buffer.Append("          .wraptext  = ").Append(WrapText)
                .Append("\n");
            buffer.Append("          .valignment= ")
                .Append(StringUtil.ToHexString(VerticalAlignment)).Append("\n");
            buffer.Append("          .justlast  = ")
                .Append(StringUtil.ToHexString(JustifyLast)).Append("\n");
            buffer.Append("          .rotation  = ")
                .Append(StringUtil.ToHexString(Rotation)).Append("\n");
            buffer.Append("    .indentionoptions= ")
                .Append(StringUtil.ToHexString(IndentionOptions)).Append("\n");
            buffer.Append("          .indent    = ")
                .Append(StringUtil.ToHexString(Indent)).Append("\n");
            buffer.Append("          .shrinktoft= ").Append(ShrinkToFit)
                .Append("\n");
            buffer.Append("          .mergecells= ").Append(MergeCells)
                .Append("\n");
            buffer.Append("          .Readngordr= ")
                .Append(StringUtil.ToHexString(ReadingOrder)).Append("\n");
            buffer.Append("          .formatflag= ")
                .Append(IsIndentNotParentFormat).Append("\n");
            buffer.Append("          .fontflag  = ")
                .Append(IsIndentNotParentFont).Append("\n");
            buffer.Append("          .prntalgnmt= ")
                .Append(IsIndentNotParentAlignment).Append("\n");
            buffer.Append("          .borderflag= ")
                .Append(IsIndentNotParentBorder).Append("\n");
            buffer.Append("          .paternflag= ")
                .Append(IsIndentNotParentPattern).Append("\n");
            buffer.Append("          .celloption= ")
                .Append(IsIndentNotParentCellOptions).Append("\n");
            buffer.Append("    .borderoptns     = ")
                .Append(StringUtil.ToHexString(BorderOptions)).Append("\n");
            buffer.Append("          .lftln     = ")
                .Append(StringUtil.ToHexString(BorderLeft)).Append("\n");
            buffer.Append("          .rgtln     = ")
                .Append(StringUtil.ToHexString(BorderRight)).Append("\n");
            buffer.Append("          .topln     = ")
                .Append(StringUtil.ToHexString(BorderTop)).Append("\n");
            buffer.Append("          .btmln     = ")
                .Append(StringUtil.ToHexString(BorderBottom)).Append("\n");
            buffer.Append("    .paleteoptns     = ")
                .Append(StringUtil.ToHexString(PaletteOptions)).Append("\n");
            buffer.Append("          .leftborder= ")
                .Append(StringUtil.ToHexString(LeftBorderPaletteIdx))
                .Append("\n");
            buffer.Append("          .rghtborder= ")
                .Append(StringUtil.ToHexString(RightBorderPaletteIdx))
                .Append("\n");
            buffer.Append("          .diag      = ")
                .Append(StringUtil.ToHexString(Diagonal)).Append("\n");
            buffer.Append("    .paleteoptn2     = ")
                .Append(StringUtil.ToHexString(AdtlPaletteOptions))
                .Append("\n");
            buffer.Append("          .topborder = ")
                .Append(StringUtil.ToHexString(TopBorderPaletteIdx))
                .Append("\n");
            buffer.Append("          .botmborder= ")
                .Append(StringUtil.ToHexString(BottomBorderPaletteIdx))
                .Append("\n");
            buffer.Append("          .adtldiag  = ")
                .Append(StringUtil.ToHexString(AdtlDiagBorderPaletteIdx)).Append("\n");
            buffer.Append("          .diaglnstyl= ")
                .Append(StringUtil.ToHexString(AdtlDiagLineStyle)).Append("\n");
            buffer.Append("          .Fillpattrn= ")
                .Append(StringUtil.ToHexString(AdtlFillPattern)).Append("\n");
            buffer.Append("    .Fillpaloptn     = ")
                .Append(StringUtil.ToHexString(FillPaletteOptions))
                .Append("\n");
            buffer.Append("          .foreground= ")
                .Append(StringUtil.ToHexString(FillForeground)).Append("\n");
            buffer.Append("          .background= ")
                .Append(StringUtil.ToHexString(FillBackground)).Append("\n");
            buffer.Append("[/EXTENDEDFORMAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(FontIndex);
            out1.WriteShort(FormatIndex);
            out1.WriteShort(CellOptions);
            out1.WriteShort(AlignmentOptions);
            out1.WriteShort(IndentionOptions);
            out1.WriteShort(BorderOptions);
            out1.WriteShort(PaletteOptions);
            out1.WriteInt(AdtlPaletteOptions);
            out1.WriteShort(FillPaletteOptions);
        }

        protected override int DataSize
        {
            get { return 20; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override int GetHashCode()
        {
            int prime = 31;
            int result = 1;
            result = prime * result + field_1_font_index;
            result = prime * result + field_2_format_index;
            result = prime * result + field_3_cell_options;
            result = prime * result + field_4_alignment_options;
            result = prime * result + field_5_indention_options;
            result = prime * result + field_6_border_options;
            result = prime * result + field_7_palette_options;
            result = prime * result + field_8_adtl_palette_options;
            result = prime * result + field_9_fill_palette_options;
            return result;
        }

        /**
         * Will consider two different records with the same
         *  contents as Equals, as the various indexes
         *  that matter are embedded in the records
         */
        public override bool Equals(Object obj)
        {
            if (this == obj)
                return true;
            if (obj == null)
                return false;
            if (obj is ExtendedFormatRecord)
            {
                ExtendedFormatRecord other = (ExtendedFormatRecord)obj;
                if (field_1_font_index != other.field_1_font_index)
                    return false;
                if (field_2_format_index != other.field_2_format_index)
                    return false;
                if (field_3_cell_options != other.field_3_cell_options)
                    return false;
                if (field_4_alignment_options != other.field_4_alignment_options)
                    return false;
                if (field_5_indention_options != other.field_5_indention_options)
                    return false;
                if (field_6_border_options != other.field_6_border_options)
                    return false;
                if (field_7_palette_options != other.field_7_palette_options)
                    return false;
                if (field_8_adtl_palette_options != other.field_8_adtl_palette_options)
                    return false;
                if (field_9_fill_palette_options != other.field_9_fill_palette_options)
                    return false;
                return true;
            }
            return false;
        }

        public int[] StateSummary
        {
            get
            {
                return new int[] { field_1_font_index, field_2_format_index, field_3_cell_options, field_4_alignment_options,
                field_5_indention_options, field_6_border_options, field_7_palette_options, field_8_adtl_palette_options, field_9_fill_palette_options };
            }
        }
    }
}
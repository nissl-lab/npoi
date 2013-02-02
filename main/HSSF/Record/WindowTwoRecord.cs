
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
     * Title:        Window Two Record
     * Description:  sheet window Settings
     * REFERENCE:  PG 422 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class WindowTwoRecord
       : StandardRecord
    {
        public const short sid = 0x23e;
        

        // bitfields
        private BitField displayFormulas = BitFieldFactory.GetInstance(0x01);
        private BitField displayGridlines = BitFieldFactory.GetInstance(0x02);
        private BitField displayRowColHeadings = BitFieldFactory.GetInstance(0x04);
        private BitField freezePanes = BitFieldFactory.GetInstance(0x08);
        private BitField displayZeros = BitFieldFactory.GetInstance(0x10);
        // if false use color in field 4 if true use default foreground for headers
        private BitField defaultHeader =  BitFieldFactory.GetInstance(0x20);   
        private BitField arabic = BitFieldFactory.GetInstance(0x40);   
        private BitField displayGuts = BitFieldFactory.GetInstance(0x80);
        private BitField freezePanesNoSplit = BitFieldFactory.GetInstance(0x100);
        private BitField selected = BitFieldFactory.GetInstance(0x200);
        private BitField active = BitFieldFactory.GetInstance(0x400);
        private BitField savedInPageBreakPreview = BitFieldFactory.GetInstance(0x800);

        // 4-7 reserved
        // end bitfields

        private short field_1_options;
        private short field_2_top_row;
        private short field_3_left_col;
        private int field_4_header_color;
        private short field_5_page_break_zoom;
        private short field_6_normal_zoom;
        private int field_7_reserved;

        public WindowTwoRecord()
        {
        }

        /**
         * Constructs a WindowTwo record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public WindowTwoRecord(RecordInputStream in1)
        {
            int size = in1.Remaining;
            field_1_options = in1.ReadShort();
            field_2_top_row = in1.ReadShort();
            field_3_left_col = in1.ReadShort();
            field_4_header_color = in1.ReadInt();
            if (size > 10)
            {
                field_5_page_break_zoom = in1.ReadShort();
                field_6_normal_zoom = in1.ReadShort();
            }
            if (size > 14)
            {   // there Is a special case of this record that has only 14 bytes...undocumented!
                field_7_reserved = in1.ReadInt();
            }
        }


        /**
         * Get the options bitmask or just use the bit Setters.
         * @return options
         */

        public short Options
        {
            get { return field_1_options; }
            set { field_1_options = value; }
        }

        // option bitfields

        /**
         * Get whether the window should Display formulas
         * @return formulas or not
         */

        public bool DisplayFormulas
        {
            get
            {
                return displayFormulas.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = displayFormulas.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Get whether the window should Display gridlines
         * @return gridlines or not
         */

        public bool DisplayGridlines
        {
            get
            {
                return displayGridlines.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = displayGridlines.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Get whether the window should Display row and column headings
         * @return headings or not
         */

        public bool DisplayRowColHeadings
        {
            get
            {
                return displayRowColHeadings.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = displayRowColHeadings.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Get whether the window should freeze panes
         * @return freeze panes or not
         */

        public bool FreezePanes
        {
            get
            {
                return freezePanes.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = freezePanes.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Get whether the window should Display zero values
         * @return zeros or not
         */

        public bool DisplayZeros
        {
            get
            {
                return displayZeros.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = displayZeros.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Get whether the window should Display a default header
         * @return header or not
         */

        public bool DefaultHeader
        {
            get
            {
                return defaultHeader.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = defaultHeader.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Is this arabic?
         * @return arabic or not
         */

        public bool Arabic
        {
            get
            {
                return arabic.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = arabic.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Get whether the outline symbols are displaed
         * @return symbols or not
         */

        public bool DisplayGuts
        {
            get
            {
                return displayGuts.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = displayGuts.SetShortBoolean(field_1_options, value); 
            }
        }

        /**
         * freeze Unsplit panes or not
         * @return freeze or not
         */

        public bool FreezePanesNoSplit
        {
            get
            {
                return freezePanesNoSplit.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = freezePanesNoSplit.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * sheet tab Is selected
         * @return selected or not
         */

        public bool IsSelected
        {
            get
            {
                return selected.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = selected.SetShortBoolean(field_1_options, value);
            }
        }

        /**
         * Is the sheet currently Displayed in the window
         * @return Displayed or not
         */

        public bool IsActive
        {
            get
            {
                return active.IsSet(field_1_options);
            }
            set 
            { 
                field_1_options = active.SetShortBoolean(field_1_options, value); 
            }
        }
        /**
         * deprecated May 2008
         * @deprecated use IsActive()
         */
        [Obsolete]
        public bool Paged
        {
            get
            {
                return IsActive;
            }
            set 
            {
                IsActive = value;
            }
        }

        /**
         * was the sheet saved in page break view
         * @return pagebreaksaved or not
         */

        public bool SavedInPageBreakPreview
        {
            get
            {
                return savedInPageBreakPreview.IsSet(field_1_options);
            }
            set 
            {
                field_1_options = savedInPageBreakPreview.SetShortBoolean(field_1_options, value);
            }
        }

        // end of bitfields.

        /**
         * Get the top row visible in the window
         * @return toprow
         */

        public short TopRow
        {
            get
            {
                return field_2_top_row;
            }
            set 
            {
                field_2_top_row = value;
            }
        }

        /**
         * Get the leftmost column Displayed in the window
         * @return leftmost
         */

        public short LeftCol
        {
            get
            {
                return field_3_left_col;
            }
            set 
            {
                field_3_left_col = value;
            }
        }

        /**
         * Get the palette index for the header color
         * @return color
         */

        public int HeaderColor
        {
            get
            {
                return field_4_header_color;
            }
            set 
            {
                field_4_header_color = value;
            }
        }

        /**
         * zoom magification in page break view
         * @return zoom
         */

        public short PageBreakZoom
        {
            get
            {
                return field_5_page_break_zoom;
            }
            set 
            {
                field_5_page_break_zoom = value; 
            }
        }

        /**
         * Get the zoom magnification in normal view
         * @return zoom
         */

        public short NormalZoom
        {
            get
            {
                return field_6_normal_zoom;
            }
            set 
            {
                field_6_normal_zoom = value;
            }
        }

        /**
         * Get the reserved bits - why would you do this?
         * @return reserved stuff -probably garbage
         */

        public int Reserved
        {
            get { return field_7_reserved; }
            set { field_7_reserved = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[WINDOW2]\n");
            buffer.Append("    .options        = ")
                .Append(StringUtil.ToHexString(Options)).Append("\n");
            buffer.Append("       .dispformulas= ").Append(DisplayFormulas)
                .Append("\n");
            buffer.Append("       .dispgridlins= ").Append(DisplayGridlines)
                .Append("\n");
            buffer.Append("       .disprcheadin= ")
                .Append(DisplayRowColHeadings).Append("\n");
            buffer.Append("       .freezepanes = ").Append(FreezePanes)
                .Append("\n");
            buffer.Append("       .Displayzeros= ").Append(DisplayZeros)
                .Append("\n");
            buffer.Append("       .defaultheadr= ").Append(DefaultHeader)
                .Append("\n");
            buffer.Append("       .arabic      = ").Append(Arabic)
                .Append("\n");
            buffer.Append("       .Displayguts = ").Append(DisplayGuts)
                .Append("\n");
            buffer.Append("       .frzpnsnosplt= ")
                .Append(FreezePanesNoSplit).Append("\n");
            buffer.Append("       .selected    = ").Append(IsSelected)
                .Append("\n");
            buffer.Append("       .active       = ").Append(IsActive)
                .Append("\n");
            buffer.Append("       .svdinpgbrkpv= ")
                .Append(SavedInPageBreakPreview).Append("\n");
            buffer.Append("    .toprow         = ")
                .Append(StringUtil.ToHexString(TopRow)).Append("\n");
            buffer.Append("    .leftcol        = ")
                .Append(StringUtil.ToHexString(LeftCol)).Append("\n");
            buffer.Append("    .headercolor    = ")
                .Append(StringUtil.ToHexString(HeaderColor)).Append("\n");
            buffer.Append("    .pagebreakzoom  = ")
                .Append(StringUtil.ToHexString(PageBreakZoom)).Append("\n");
            buffer.Append("    .normalzoom     = ")
                .Append(StringUtil.ToHexString(NormalZoom)).Append("\n");
            buffer.Append("    .reserved       = ")
                .Append(StringUtil.ToHexString(Reserved)).Append("\n");
            buffer.Append("[/WINDOW2]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Options);
            out1.WriteShort(TopRow);
            out1.WriteShort(LeftCol);
            out1.WriteInt(HeaderColor);
            out1.WriteShort(PageBreakZoom);
            out1.WriteShort(NormalZoom);
            out1.WriteInt(Reserved);
        }

        protected override int DataSize
        {
            get
            {
                return 18;
            }
        }
        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            WindowTwoRecord rec = new WindowTwoRecord();
            rec.field_1_options = field_1_options;
            rec.field_2_top_row = field_2_top_row;
            rec.field_3_left_col = field_3_left_col;
            rec.field_4_header_color = field_4_header_color;
            rec.field_5_page_break_zoom = field_5_page_break_zoom;
            rec.field_6_normal_zoom = field_6_normal_zoom;
            rec.field_7_reserved = field_7_reserved;
            return rec;
        }
    }
}

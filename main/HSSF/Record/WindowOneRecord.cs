
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
     * Title:        Window1 Record
     * Description:  Stores the attributes of the workbook window.  This Is basically
     *               so the gui knows how big to make the window holding the spReadsheet
     *               document.
     * REFERENCE:  PG 421 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class WindowOneRecord
       : StandardRecord
    {
        public const short sid = 0x3d;

        // our variable names stolen from old TV Sets.
        private short field_1_h_hold;                  // horizontal position
        private short field_2_v_hold;                  // vertical position
        private short field_3_width;
        private short field_4_height;
        private short field_5_options;
        static private BitField hidden =
            BitFieldFactory.GetInstance(0x01);                                        // Is this window Is hidden
        static private BitField iconic =
            BitFieldFactory.GetInstance(0x02);                                        // Is this window Is an icon
        static private BitField reserved = BitFieldFactory.GetInstance(0x04);   // reserved
        static private BitField hscroll =
            BitFieldFactory.GetInstance(0x08);                                        // Display horizontal scrollbar
        static private BitField vscroll =
            BitFieldFactory.GetInstance(0x10);                                        // Display vertical scrollbar
        static private BitField tabs =
            BitFieldFactory.GetInstance(0x20);                                        // Display tabs at the bottom

        // all the rest are "reserved"
        private int field_6_active_sheet;
        private int field_7_first_visible_tab;
        private short field_8_num_selected_tabs;
        private short field_9_tab_width_ratio;

        public WindowOneRecord()
        {
        }

        /**
         * Constructs a WindowOne record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public WindowOneRecord(RecordInputStream in1)
        {
            field_1_h_hold = in1.ReadShort();
            field_2_v_hold = in1.ReadShort();
            field_3_width = in1.ReadShort();
            field_4_height = in1.ReadShort();
            field_5_options = in1.ReadShort();
            field_6_active_sheet = in1.ReadShort();
            field_7_first_visible_tab = in1.ReadShort();
            field_8_num_selected_tabs = in1.ReadShort();
            field_9_tab_width_ratio = in1.ReadShort();
        }

        /**
         * Get the horizontal position of the window (in 1/20ths of a point)
         * @return h - horizontal location
         */

        public short HorizontalHold
        {
            get { return field_1_h_hold; }
            set { field_1_h_hold = value; }
        }

        /**
         * Get the vertical position of the window (in 1/20ths of a point)
         * @return v - vertical location
         */

        public short VerticalHold
        {
            get { return field_2_v_hold; }
            set { field_2_v_hold = value; }
        }

        /**
         * Get the width of the window
         * @return width
         */

        public short Width
        {
            get { return field_3_width; }
            set { field_3_width = value; }
        }

        /**
         * Get the height of the window
         * @return height
         */

        public short Height
        {
            get { return field_4_height; }
            set { field_4_height = value; }
        }

        /**
         * Get the options bitmask (see bit Setters)
         *
         * @return o - the bitmask
         */

        public short Options
        {
            get { return field_5_options; }
            set { field_5_options = value; }
        }

        // bitfields for options

        /**
         * Get whether the window Is hidden or not
         * @return Ishidden or not
         */

        public bool Hidden
        {
            get { return hidden.IsSet(field_5_options); }
            set { field_5_options = hidden.SetShortBoolean(field_5_options, value); }
        }

        /**
         * Get whether the window has been iconized or not
         * @return iconize  or not
         */

        public bool Iconic
        {
            get { return iconic.IsSet(field_5_options); }
            set { field_5_options = iconic.SetShortBoolean(field_5_options, value); }
        }

        /**
         * Get whether to Display the horizontal scrollbar or not
         * @return Display or not
         */

        public bool DisplayHorizontalScrollbar
        {
            get{return hscroll.IsSet(field_5_options);}
            set { field_5_options = hscroll.SetShortBoolean(field_5_options, value); }
        }

        /**
         * Get whether to Display the vertical scrollbar or not
         * @return Display or not
         */

        public bool DisplayVerticalScrollbar
        {
            get { return vscroll.IsSet(field_5_options); }
            set { field_5_options = vscroll.SetShortBoolean(field_5_options, value); }
        }

        /**
         * Get whether to Display the tabs or not
         * @return Display or not
         */

        public bool DisplayTabs
        {
            get { return tabs.IsSet(field_5_options); }
            set { field_5_options = tabs.SetShortBoolean(field_5_options, value); }
        }

        // end options bitfields


        /**
         * @return the index of the currently Displayed sheet 
         */
        public int ActiveSheetIndex
        {
            get { return field_6_active_sheet; }
            set { field_6_active_sheet = value; }
        }
        /**
         * deprecated May 2008
         * @deprecated - Misleading name - use GetActiveSheetIndex() 
         */
        [Obsolete]
        public short SelectedTab
        {
            get { return (short)ActiveSheetIndex; }
            set { ActiveSheetIndex = value; }
        }

        /**
         * @return the first visible sheet in the worksheet tab-bar. 
         * I.E. the scroll position of the tab-bar.
         */
        public int FirstVisibleTab
        {
            get { return field_7_first_visible_tab; }
            set { field_7_first_visible_tab = value; }
        }
        /**
         * deprecated May 2008
         * @deprecated - Misleading name - use GetFirstVisibleTab() 
         */
        [Obsolete]
        public short DisplayedTab
        {
            get { return (short)FirstVisibleTab; }
            set { FirstVisibleTab=value; }
        }

        /**
         * Get the number of selected tabs
         * @return number of tabs
         */

        public short NumSelectedTabs
        {
            get { return field_8_num_selected_tabs; }
            set { field_8_num_selected_tabs = value; }
        }

        /**
         * ratio of the width of the tabs to the horizontal scrollbar
         * @return ratio
         */

        public short TabWidthRatio
        {
            get{return field_9_tab_width_ratio;}
            set { field_9_tab_width_ratio = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[WINDOW1]\n");
            buffer.Append("    .h_hold          = ")
                .Append(StringUtil.ToHexString(HorizontalHold)).Append("\n");
            buffer.Append("    .v_hold          = ")
                .Append(StringUtil.ToHexString(VerticalHold)).Append("\n");
            buffer.Append("    .width           = ")
                .Append(StringUtil.ToHexString(Width)).Append("\n");
            buffer.Append("    .height          = ")
                .Append(StringUtil.ToHexString(Height)).Append("\n");
            buffer.Append("    .options         = ")
                .Append(StringUtil.ToHexString(Options)).Append("\n");
            buffer.Append("        .hidden      = ").Append(Hidden)
                .Append("\n");
            buffer.Append("        .iconic      = ").Append(Iconic)
                .Append("\n");
            buffer.Append("        .hscroll     = ")
                .Append(DisplayHorizontalScrollbar).Append("\n");
            buffer.Append("        .vscroll     = ")
                .Append(DisplayVerticalScrollbar).Append("\n");
            buffer.Append("        .tabs        = ").Append(DisplayTabs)
                .Append("\n");
            buffer.Append("    .activeSheet     = ")
                .Append(StringUtil.ToHexString(ActiveSheetIndex)).Append("\n");
            buffer.Append("    .firstVisibleTab    = ")
                .Append(StringUtil.ToHexString(FirstVisibleTab)).Append("\n");
            buffer.Append("    .numselectedtabs = ")
                .Append(StringUtil.ToHexString(NumSelectedTabs)).Append("\n");
            buffer.Append("    .tabwidthratio   = ")
                .Append(StringUtil.ToHexString(TabWidthRatio)).Append("\n");
            buffer.Append("[/WINDOW1]\n");
            return buffer.ToString();
        }

        
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(HorizontalHold);
            out1.WriteShort(VerticalHold);
            out1.WriteShort(Width);
            out1.WriteShort(Height);
            out1.WriteShort(Options);
            out1.WriteShort(ActiveSheetIndex);
            out1.WriteShort(FirstVisibleTab);
            out1.WriteShort(NumSelectedTabs);
            out1.WriteShort(TabWidthRatio);
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
    }
}

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
     * Title:        Row Record
     * Description:  stores the row information for the sheet. 
     * REFERENCE:  PG 379 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */
    public class RowRecord : StandardRecord, IComparable
    {
        public const short sid = 0x208;
        public const int ENCODED_SIZE = 20;

        private const int OPTION_BITS_ALWAYS_SET = 0x0100;
        //private const int DEFAULT_HEIGHT_BIT = 0x8000;

        /** The maximum row number that excel can handle (zero based) ie 65536 rows Is
         *  max number of rows.
         */
        [Obsolete]
        public const int MAX_ROW_NUMBER = 65535;

        private int field_1_row_number;
        private int field_2_first_col;
        private int field_3_last_col;   // plus 1
        private short field_4_height;
        private short field_5_optimize;   // hint field for gui, can/should be Set to zero

        // for generated sheets.
        private short field_6_reserved;
        /** 16 bit options flags */
        private int field_7_option_flags;
        private static BitField outlineLevel = BitFieldFactory.GetInstance(0x07);

        // bit 3 reserved
        private static BitField colapsed = BitFieldFactory.GetInstance(0x10);
        private static BitField zeroHeight = BitFieldFactory.GetInstance(0x20);
        private static BitField badFontHeight = BitFieldFactory.GetInstance(0x40);
        private static BitField formatted = BitFieldFactory.GetInstance(0x80);

    private int field_8_option_flags;   // only if isFormatted
    private static BitField          xfIndex       = BitFieldFactory.GetInstance(0xFFF);
    private static BitField          topBorder     = BitFieldFactory.GetInstance(0x1000);
    private static BitField          bottomBorder  = BitFieldFactory.GetInstance(0x2000);
    private static BitField          phoeneticGuide  = BitFieldFactory.GetInstance(0x4000);

        public RowRecord(int rowNumber)
        {
            field_1_row_number = rowNumber;
            //field_2_first_col = -1;
            //field_3_last_col = -1;
            field_4_height = (short)0x00FF;
            field_5_optimize = (short)0;
            field_6_reserved = (short)0;
            field_7_option_flags = OPTION_BITS_ALWAYS_SET; // seems necessary for outlining

            field_8_option_flags = (short)0xf;
            SetEmpty();
        }

        /**
         * Constructs a Row record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public RowRecord(RecordInputStream in1)
        {
            field_1_row_number = in1.ReadUShort();
            field_2_first_col = in1.ReadShort();
            field_3_last_col = in1.ReadShort();
            field_4_height = in1.ReadShort();
            field_5_optimize = in1.ReadShort();
            field_6_reserved = in1.ReadShort();
            field_7_option_flags = in1.ReadShort();
            field_8_option_flags = in1.ReadShort();
        }

        public void SetEmpty()
        { 
            field_2_first_col = 0;
            field_3_last_col = 0;            
        }
        /**
         * Get the logical row number for this row (0 based index)
         * @return row - the row number
         */
        public bool IsEmpty
        {
            get
            {
                return (field_2_first_col | field_3_last_col) == 0;
            }
        }
        //public short RowNumber
        public int RowNumber
        {
            get
            {
                return field_1_row_number;
            }
            set { field_1_row_number = value; }
        }

        /**
         * Get the logical col number for the first cell this row (0 based index)
         * @return col - the col number
         */

        public int FirstCol
        {
            get{return field_2_first_col;}
            set { field_2_first_col = value; }
        }

        /**
         * Get the logical col number for the last cell this row plus one (0 based index)
         * @return col - the last col number + 1
         */

        public int LastCol
        {
            get { return field_3_last_col; }
            set { field_3_last_col = value; }
        }

        /**
         * Get the height of the row
         * @return height of the row
         */

        public short Height
        {
            get{return field_4_height;}
            set { field_4_height = value; }
        }

        /**
         * Get whether to optimize or not (Set to 0)
         * @return optimize (Set to 0)
         */

        public short Optimize
        {
            get
            {
                return field_5_optimize;
            }
            set { field_5_optimize = value; }
        }

        /**
         * Gets the option bitmask.  (use the individual bit Setters that refer to this
         * method)
         * @return options - the bitmask
         */

        public short OptionFlags
        {
            get
            {
                return (short)field_7_option_flags;
            }
            set { field_7_option_flags = value | (short)OPTION_BITS_ALWAYS_SET; }
        }

        // option bitfields

        /**
         * Get the outline level of this row
         * @return ol - the outline level
         * @see #GetOptionFlags()
         */

        public short OutlineLevel
        {
            get { return (short)outlineLevel.GetValue(field_7_option_flags); }
            set { field_7_option_flags = outlineLevel.SetValue(field_7_option_flags, value); }
        }

        /**
         * Get whether or not to colapse this row
         * @return c - colapse or not
         * @see #GetOptionFlags()
         */

        public bool Colapsed
        {
            get
            {
                return (colapsed.IsSet(field_7_option_flags));
            }
            set 
            { 
                field_7_option_flags = colapsed.SetBoolean(field_7_option_flags, value); 
            }
        }

        /**
         * Get whether or not to Display this row with 0 height
         * @return - z height is zero or not.
         * @see #GetOptionFlags()
         */

        public bool ZeroHeight
        {
            get
            {
                return zeroHeight.IsSet(field_7_option_flags);
            }
            set 
            { 
                field_7_option_flags = zeroHeight.SetBoolean(field_7_option_flags, value); 
            }
        }

        /**
         * Get whether the font and row height are not compatible
         * @return - f -true if they aren't compatible (damn not logic)
         * @see #GetOptionFlags()
         */

        public bool BadFontHeight
        {
            get
            {
                return badFontHeight.IsSet(field_7_option_flags);
            }
            set 
            { 
                field_7_option_flags = badFontHeight.SetBoolean(field_7_option_flags, value); 
            }
        }

        /**
         * Get whether the row has been formatted (even if its got all blank cells)
         * @return formatted or not
         * @see #GetOptionFlags()
         */

        public bool Formatted
        {
            get
            {
                return formatted.IsSet(field_7_option_flags);
            }
            set { field_7_option_flags = formatted.SetBoolean(field_7_option_flags, value); }
        }

        // end bitfields
        public short OptionFlags2
        {
            get { return (short)this.field_8_option_flags; }
        }
        /**
         * if the row is formatted then this is the index to the extended format record
         * @see org.apache.poi.hssf.record.ExtendedFormatRecord
         * @return index to the XF record or bogus value (undefined) if Isn't formatted
         */

        public short XFIndex
        {
            get { return xfIndex.GetShortValue((short)field_8_option_flags); }
            set { field_8_option_flags = xfIndex.SetValue(field_8_option_flags, value); }
        }
        /**
         * bit that specifies whether any cell in the row has a thick top border, or any
         * cell in the row directly above the current row has a thick bottom border.
         * @param f has thick top border
         */
        public bool TopBorder
        {
            get { return topBorder.IsSet(field_8_option_flags); }
            set { field_8_option_flags = topBorder.SetBoolean(field_8_option_flags, value); }
        }
        /**
         * A bit that specifies whether any cell in the row has a medium or thick
         * bottom border, or any cell in the row directly below the current row has
         * a medium or thick top border.
         * @param f has thick bottom border
         */
        public bool BottomBorder
        {
            get { return bottomBorder.IsSet(field_8_option_flags); }
            set { field_8_option_flags = bottomBorder.SetBoolean(field_8_option_flags, value); }
        }
        /**
         * A bit that specifies whether the phonetic guide feature is enabled for
         * any cell in this row.
         * @param f use phoenetic guide
         */
        public bool PhoeneticGuide
        {
            get { return phoeneticGuide.IsSet(field_8_option_flags); }
            set { field_8_option_flags = phoeneticGuide.SetBoolean(field_8_option_flags, value); }
        }
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ROW]\n");
            buffer.Append("    .rownumber      = ").Append(StringUtil.ToHexString(RowNumber)).Append("\n");
            buffer.Append("    .firstcol       = ").Append(StringUtil.ToHexString(FirstCol)).Append("\n");
            buffer.Append("    .lastcol        = ").Append(StringUtil.ToHexString(LastCol)).Append("\n");
            buffer.Append("    .height         = ").Append(StringUtil.ToHexString(Height)).Append("\n");
            buffer.Append("    .optimize       = ").Append(StringUtil.ToHexString(Optimize)).Append("\n");
            buffer.Append("    .reserved       = ").Append(StringUtil.ToHexString(field_6_reserved)).Append("\n");
            buffer.Append("    .optionflags    = ").Append(StringUtil.ToHexString(OptionFlags)).Append("\n");
            buffer.Append("        .outlinelvl = ").Append(StringUtil.ToHexString(OutlineLevel)).Append("\n");
            buffer.Append("        .colapsed   = ").Append(Colapsed).Append("\n");
            buffer.Append("        .zeroheight = ").Append(ZeroHeight).Append("\n");
            buffer.Append("        .badfontheig= ").Append(BadFontHeight).Append("\n");
            buffer.Append("        .formatted  = ").Append(Formatted).Append("\n");
            buffer.Append("    .optionsflags2  = ").Append(StringUtil.ToHexString(OptionFlags2)).Append("\n");
            buffer.Append("        .xFindex       = ").Append(StringUtil.ToHexString(XFIndex)).Append("\n");
            buffer.Append("        .topBorder     = ").Append(TopBorder).Append("\n");
            buffer.Append("        .bottomBorder  = ").Append(BottomBorder).Append("\n");
            buffer.Append("        .phoeneticGuide= ").Append(PhoeneticGuide).Append("\n");
            buffer.Append("[/ROW]\n");
            return buffer.ToString();
        }


        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(RowNumber);
            out1.WriteShort(FirstCol == -1 ? (short)0 : FirstCol);
            out1.WriteShort(LastCol == -1 ? (short)0 : LastCol);
            out1.WriteShort(Height);
            out1.WriteShort(Optimize);
            out1.WriteShort(field_6_reserved);
            out1.WriteShort(OptionFlags);
            out1.WriteShort(OptionFlags2);
        }
        protected override int DataSize
        {
            get
            {
                return ENCODED_SIZE - 4;
            }
        }
        public override int RecordSize
        {
            get { return 20; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public int CompareTo(Object obj)
        {
            RowRecord loc = (RowRecord)obj;

            if (this.RowNumber == loc.RowNumber)
            {
                return 0;
            }
            if (this.RowNumber < loc.RowNumber)
            {
                return -1;
            }
            if (this.RowNumber > loc.RowNumber)
            {
                return 1;
            }
            return -1;
        }

        public override bool Equals(Object obj)
        {
            if (!(obj is RowRecord))
            {
                return false;
            }
            RowRecord loc = (RowRecord)obj;

            if (this.RowNumber == loc.RowNumber)
            {
                return true;
            }
            return false;
        }

        public override int GetHashCode ()
        {
            return RowNumber;
        }

        public override Object Clone()
        {
            RowRecord rec = new RowRecord(field_1_row_number);
            rec.field_2_first_col = field_2_first_col;
            rec.field_3_last_col = field_3_last_col;
            rec.field_4_height = field_4_height;
            rec.field_5_optimize = field_5_optimize;
            rec.field_6_reserved = field_6_reserved;
            rec.field_7_option_flags = field_7_option_flags;
            rec.field_8_option_flags = field_8_option_flags;
            return rec;
        }
    }
}
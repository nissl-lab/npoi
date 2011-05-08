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
    using System.Collections;
    using System;
    using System.IO;
    using System.Text;
    using NPOI.Util;

    /**
     * Title: COLINFO Record<p/>
     * Description:  Defines with width and formatting for a range of columns<p/>
     * REFERENCE:  PG 293 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)<p/>
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */
    public class ColumnInfoRecord : Record
    {
        public const short sid = 0x7d;
        private int field_1_first_col;
        private int field_2_last_col;
        private int field_3_col_width;
        private int field_4_xf_index;
        private int field_5_options;
        private static BitField hidden = BitFieldFactory.GetInstance(0x01);
        private static BitField outlevel = BitFieldFactory.GetInstance(0x0700);
        private static BitField collapsed = BitFieldFactory.GetInstance(0x1000);
        // Excel seems Write values 2, 10, and 260, even though spec says "must be zero"
        private int field_6_reserved;

        public ColumnInfoRecord()
        {
            this.ColumnWidth = 2275;
            field_5_options = 2;
            field_4_xf_index = 0x0f;
            field_6_reserved = 2; // seems to be the most common value
        }

        /**
         * Constructs a ColumnInfo record and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public ColumnInfoRecord(RecordInputStream in1)
        {
            field_1_first_col = in1.ReadUShort();
            field_2_last_col = in1.ReadUShort();
            field_3_col_width = in1.ReadShort();
            field_4_xf_index = in1.ReadShort();
            field_5_options = in1.ReadShort();
            switch (in1.Remaining)
            {
                case 2: // usual case
                    field_6_reserved = in1.ReadShort();
                    break;
                case 1:
                    // often COLINFO Gets encoded 1 byte short
                    // shouldn't matter because this field Is Unused
                    field_6_reserved = in1.ReadByte();
                    break;
                default:
                    throw new Exception("Unusual record size remaining=(" + in1.Remaining + ")");
            }
        }
        /**
         * @return true if the format, options and column width match
         */
        public bool FormatMatches(ColumnInfoRecord other)
        {
            if (field_4_xf_index != other.field_4_xf_index)
            {
                return false;
            }
            if (field_5_options != other.field_5_options)
            {
                return false;
            }
            if (field_3_col_width != other.field_3_col_width)
            {
                return false;
            }
            return true;
        }

        /**
         * Get the first column this record defines formatting info for
         * @return the first column index (0-based)
         */

        public int FirstColumn
        {
            get{return field_1_first_col;}
            set { field_1_first_col = value; }
        }

        /**
         * Get the last column this record defines formatting info for
         * @return the last column index (0-based)
         */

        public int LastColumn
        {
            get { return field_2_last_col; }
            set { field_2_last_col = value; }
        }

        /**
         * Get the columns' width in 1/256 of a Char width
         * @return column width
         */

        public int ColumnWidth
        {
            get
            {
                return field_3_col_width;
            }
            set { field_3_col_width = value; }
        }

        /**
         * Get the columns' default format info
         * @return the extended format index
         * @see org.apache.poi.hssf.record.ExtendedFormatRecord
         */

        public int XFIndex
        {
            get { return field_4_xf_index; }
            set { field_4_xf_index = value; }
        }

        /**
         * Get the options bitfield - use the bitSetters instead
         * @return the bitfield raw value
         */

        public int Options
        {
            get { return field_5_options; }
            set { field_5_options = value; }
        }

        // start options bitfield

        /**
         * Get whether or not these cells are hidden
         * @return whether the cells are hidden.
         * @see #SetOptions(short)
         */

        public bool IsHidden
        {
            get { return hidden.IsSet(field_5_options); }
            set { field_5_options = hidden.SetBoolean(field_5_options, value); }
        }

        /**
         * Get the outline level for the cells
         * @see #SetOptions(short)
         * @return outline level for the cells
         */

        public int OutlineLevel
        {
            get { return outlevel.GetValue(field_5_options); }
            set { field_5_options = outlevel.SetValue(field_5_options, value); }
        }

        /**
         * Get whether the cells are collapsed
         * @return wether the cells are collapsed
         * @see #SetOptions(short)
         */

        public bool IsCollapsed
        {
            get { return collapsed.IsSet(field_5_options); }
            set
            {
                field_5_options = collapsed.SetBoolean(field_5_options,
                                                    value);
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public bool ContainsColumn(int columnIndex)
        {
            return field_1_first_col <= columnIndex && columnIndex <= field_2_last_col;
        }
        public bool IsAdjacentBefore(ColumnInfoRecord other)
        {
            return field_2_last_col == other.field_1_first_col - 1;
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutUShort(data, 2 + offset, 12);
            LittleEndian.PutUShort(data, 4 + offset, FirstColumn);
            LittleEndian.PutUShort(data, 6 + offset, LastColumn);
            LittleEndian.PutUShort(data, 8 + offset, ColumnWidth);
            LittleEndian.PutUShort(data, 10 + offset, XFIndex);
            LittleEndian.PutUShort(data, 12 + offset, field_5_options);
            LittleEndian.PutUShort(data, 14 + offset, field_6_reserved);
            return RecordSize;
        }

        public override int RecordSize
        {
            get { return 16; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[COLINFO]\n");
            buffer.Append("colfirst       = ").Append(FirstColumn)
                .Append("\n");
            buffer.Append("collast        = ").Append(LastColumn)
                .Append("\n");
            buffer.Append("colwidth       = ").Append(ColumnWidth)
                .Append("\n");
            buffer.Append("xFindex        = ").Append(XFIndex).Append("\n");
            buffer.Append("options        = ").Append(Options).Append("\n");
            buffer.Append("  hidden       = ").Append(IsHidden).Append("\n");
            buffer.Append("  olevel       = ").Append(OutlineLevel)
                .Append("\n");
            buffer.Append("  collapsed    = ").Append(IsCollapsed)
                .Append("\n");
            buffer.Append("[/COLINFO]\n");
            return buffer.ToString();
        }

        public override Object Clone()
        {
            ColumnInfoRecord rec = new ColumnInfoRecord();
            rec.field_1_first_col = field_1_first_col;
            rec.field_2_last_col = field_2_last_col;
            rec.field_3_col_width = field_3_col_width;
            rec.field_4_xf_index = field_4_xf_index;
            rec.field_5_options = field_5_options;
            rec.field_6_reserved = field_6_reserved;
            return rec;
        }
    }
}
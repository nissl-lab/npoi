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
     * Title: COLINFO Record<p/>
     * Description:  Defines with width and formatting for a range of columns<p/>
     * REFERENCE:  PG 293 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)<p/>
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */
    public class ColumnInfoRecord : StandardRecord
    {
        public const short sid = 0x7d;
        private int _first_col;
        private int _last_col;
        private int _col_width;
        private int _xf_index;
        private int _options;
        private static BitField hidden = BitFieldFactory.GetInstance(0x01);
        private static BitField outlevel = BitFieldFactory.GetInstance(0x0700);
        private static BitField collapsed = BitFieldFactory.GetInstance(0x1000);
        // Excel seems Write values 2, 10, and 260, even though spec says "must be zero"
        private int field_6_reserved;

        public ColumnInfoRecord()
        {
            this.ColumnWidth = 2275;
            _options = 2;
            _xf_index = 0x0f;
            field_6_reserved = 2; // seems to be the most common value
        }

        /**
         * Constructs a ColumnInfo record and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public ColumnInfoRecord(RecordInputStream in1)
        {
            _first_col = in1.ReadUShort();
            _last_col = in1.ReadUShort();
            _col_width = in1.ReadUShort();
            _xf_index = in1.ReadUShort();
            _options = in1.ReadUShort();
            switch (in1.Remaining)
            {
                case 2: // usual case
                    field_6_reserved = in1.ReadUShort();
                    break;
                case 1:
                    // often COLINFO Gets encoded 1 byte short
                    // shouldn't matter because this field Is Unused
                    field_6_reserved = in1.ReadByte();
                    break;
                case 0:
                    // According to bugzilla 48332,
                    // "SoftArtisans OfficeWriter for Excel" totally skips field 6
                    // Excel seems to be OK with this, and assumes zero.
                    field_6_reserved = 0;
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
            if (_xf_index != other._xf_index)
            {
                return false;
            }
            if (_options != other._options)
            {
                return false;
            }
            if (_col_width != other._col_width)
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
            get{return _first_col;}
            set { _first_col = value; }
        }

        /**
         * Get the last column this record defines formatting info for
         * @return the last column index (0-based)
         */

        public int LastColumn
        {
            get { return _last_col; }
            set { _last_col = value; }
        }

        /**
         * Get the columns' width in 1/256 of a Char width
         * @return column width
         */

        public int ColumnWidth
        {
            get
            {
                return _col_width;
            }
            set { _col_width = value; }
        }

        /**
         * Get the columns' default format info
         * @return the extended format index
         * @see org.apache.poi.hssf.record.ExtendedFormatRecord
         */

        public int XFIndex
        {
            get { return _xf_index; }
            set { _xf_index = value; }
        }

        /**
         * Get the options bitfield - use the bitSetters instead
         * @return the bitfield raw value
         */

        public int Options
        {
            get { return _options; }
            set { _options = value; }
        }

        // start options bitfield

        /**
         * Get whether or not these cells are hidden
         * @return whether the cells are hidden.
         * @see #SetOptions(short)
         */

        public bool IsHidden
        {
            get { return hidden.IsSet(_options); }
            set { _options = hidden.SetBoolean(_options, value); }
        }

        /**
         * Get the outline level for the cells
         * @see #SetOptions(short)
         * @return outline level for the cells
         */

        public int OutlineLevel
        {
            get { return outlevel.GetValue(_options); }
            set { _options = outlevel.SetValue(_options, value); }
        }

        /**
         * Get whether the cells are collapsed
         * @return wether the cells are collapsed
         * @see #SetOptions(short)
         */

        public bool IsCollapsed
        {
            get { return collapsed.IsSet(_options); }
            set
            {
                _options = collapsed.SetBoolean(_options,
                                                    value);
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public bool ContainsColumn(int columnIndex)
        {
            return _first_col <= columnIndex && columnIndex <= _last_col;
        }
        public bool IsAdjacentBefore(ColumnInfoRecord other)
        {
            return _last_col == other._first_col - 1;
        }

        protected override int DataSize
        {
            get { return 12; }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(FirstColumn);
            out1.WriteShort(LastColumn);
            out1.WriteShort(ColumnWidth);
            out1.WriteShort(XFIndex);
            out1.WriteShort(_options);
            out1.WriteShort(field_6_reserved);
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
            rec._first_col = _first_col;
            rec._last_col = _last_col;
            rec._col_width = _col_width;
            rec._xf_index = _xf_index;
            rec._options = _options;
            rec.field_6_reserved = field_6_reserved;
            return rec;
        }

        
    }
}
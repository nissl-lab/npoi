
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
     * Title:        Guts Record 
     * Description:  Row/column gutter sizes 
     * REFERENCE:  PG 320 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class GutsRecord
       : StandardRecord
    {
        public const short sid = 0x80;
        private short field_1_left_row_gutter;   // size of the row gutter to the left of the rows
        private short field_2_top_col_gutter;    // size of the column gutter above the columns
        private short field_3_row_level_max;     // maximum outline level for row gutters
        private short field_4_col_level_max;     // maximum outline level for column gutters

        public GutsRecord()
        {
        }

        /**
         * Constructs a Guts record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public GutsRecord(RecordInputStream in1)
        {
            field_1_left_row_gutter = in1.ReadShort();
            field_2_top_col_gutter = in1.ReadShort();
            field_3_row_level_max = in1.ReadShort();
            field_4_col_level_max = in1.ReadShort();
        }

        /**
         * Get the size of the gutter that appears at the left of the rows
         *
         * @return gutter size in screen Units
         */

        public short LeftRowGutter
        {
            get
            {
                return field_1_left_row_gutter;
            }
            set { field_1_left_row_gutter = value; }
        }

        /**
         * Get the size of the gutter that appears at the above the columns
         *
         * @return gutter size in screen Units
         */

        public short TopColGutter
        {
            get
            {
                return field_2_top_col_gutter;
            }
            set { field_2_top_col_gutter = value; }
        }

        /**
         * Get the maximum outline level for the row gutter.
         *
         * @return maximum outline level
         */

        public short RowLevelMax
        {
            get { return field_3_row_level_max; }
            set { field_3_row_level_max = value; }
        }

        /**
         * Get the maximum outline level for the col gutter.
         *
         * @return maximum outline level
         */

        public short ColLevelMax
        {
            get { return field_4_col_level_max; }
            set { field_4_col_level_max = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[GUTS]\n");
            buffer.Append("    .leftgutter     = ")
                .Append(StringUtil.ToHexString(LeftRowGutter)).Append("\n");
            buffer.Append("    .topgutter      = ")
                .Append(StringUtil.ToHexString(TopColGutter)).Append("\n");
            buffer.Append("    .rowlevelmax    = ")
                .Append(StringUtil.ToHexString(RowLevelMax)).Append("\n");
            buffer.Append("    .collevelmax    = ")
                .Append(StringUtil.ToHexString(ColLevelMax)).Append("\n");
            buffer.Append("[/GUTS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(LeftRowGutter);
            out1.WriteShort(TopColGutter);
            out1.WriteShort(RowLevelMax);
            out1.WriteShort(ColLevelMax);
        }

        protected override int DataSize
        {
            get
            {
                return 8;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            GutsRecord rec = new GutsRecord();
            rec.field_1_left_row_gutter = field_1_left_row_gutter;
            rec.field_2_top_col_gutter = field_2_top_col_gutter;
            rec.field_3_row_level_max = field_3_row_level_max;
            rec.field_4_col_level_max = field_4_col_level_max;
            return rec;
        }
    }
}
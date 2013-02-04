
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
    using NPOI.Util;
    using System;
    using System.Text;


    /**
     * Title:        Dimensions Record
     * Description:  provides the minumum and maximum bounds
     *               of a sheet.
     * REFERENCE:  PG 303 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class DimensionsRecord
       : StandardRecord
    {
        public const short sid = 0x200;
        private int field_1_first_row;
        private int field_2_last_row;   // plus 1
        private int field_3_first_col;
        private int field_4_last_col;
        private short field_5_zero;       // must be 0 (reserved)

        public DimensionsRecord()
        {
        }

        /**
         * Constructs a Dimensions record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public DimensionsRecord(RecordInputStream in1)
        {
            field_1_first_row = in1.ReadInt();
            field_2_last_row = in1.ReadInt();
            field_3_first_col = in1.ReadShort();
            field_4_last_col = in1.ReadShort();
            field_5_zero = in1.ReadShort();
        }

        /**
         * Get the first row number for the sheet
         * @return row - first row on the sheet
         */

        public int FirstRow
        {
            get{return field_1_first_row;}
            set { field_1_first_row = value; }
        }

        /**
         * Get the last row number for the sheet
         * @return row - last row on the sheet
         */

        public int LastRow
        {
            get { return field_2_last_row; }
            set { field_2_last_row = value; }
        }

        /**
         * Get the first column number for the sheet
         * @return column - first column on the sheet
         */

        public int FirstCol
        {
            get { return field_3_first_col; }
            set { field_3_first_col = value; }
        }

        /**
         * Get the last col number for the sheet
         * @return column - last column on the sheet
         */

        public int LastCol
        {
            get { return field_4_last_col; }
            set { field_4_last_col = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DIMENSIONS]\n");
            buffer.Append("    .firstrow       = ")
                .Append(StringUtil.ToHexString(FirstRow)).Append("\n");
            buffer.Append("    .lastrow        = ")
                .Append(StringUtil.ToHexString(LastRow)).Append("\n");
            buffer.Append("    .firstcol       = ")
                .Append(StringUtil.ToHexString(FirstCol)).Append("\n");
            buffer.Append("    .lastcol        = ")
                .Append(StringUtil.ToHexString(LastCol)).Append("\n");
            buffer.Append("    .zero           = ")
                .Append(StringUtil.ToHexString(field_5_zero)).Append("\n");
            buffer.Append("[/DIMENSIONS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(FirstRow);
            out1.WriteInt(LastRow);
            out1.WriteShort(FirstCol);
            out1.WriteShort(LastCol);
            out1.WriteShort(( short ) 0);
        }

        protected override int DataSize
        {
            get { return 14; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            DimensionsRecord rec = new DimensionsRecord();
            rec.field_1_first_row = field_1_first_row;
            rec.field_2_last_row = field_2_last_row;
            rec.field_3_first_col = field_3_first_col;
            rec.field_4_last_col = field_4_last_col;
            rec.field_5_zero = field_5_zero;
            return rec;
        }
    }
}
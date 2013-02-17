
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
    using System.Text;
    using NPOI.Util;
    using System;


    /**
     * Title:        Default Column Width Record
     * Description:  Specifies the default width for columns that have no specific
     *               width Set.
     * REFERENCE:  PG 302 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class DefaultColWidthRecord : StandardRecord
    {
        public const short sid = 0x55;
        private int field_1_col_width;
        /**
     *  The default column width is 8 characters
     */
        public const int DEFAULT_COLUMN_WIDTH = 0x0008;
        public DefaultColWidthRecord()
        {
            field_1_col_width = DEFAULT_COLUMN_WIDTH;
        }

        /**
         * Constructs a DefaultColumnWidth record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public DefaultColWidthRecord(RecordInputStream in1)
        {
            field_1_col_width = in1.ReadUShort();
        }


        /**
         * Get the default column width
         * @return defaultwidth for columns
         */

        public int ColWidth
        {
            get
            {
                return field_1_col_width;
            }
            set 
            {
                field_1_col_width = value;
            }
        }

        internal int offsetForFilePointer;  //used for defcolwidth position of IndexRecord 

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DEFAULTCOLWIDTH]\n");
            buffer.Append("    .colwidth      = ")
                .Append(StringUtil.ToHexString(ColWidth)).Append("\n");
            buffer.Append("[/DEFAULTCOLWIDTH]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(ColWidth);
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            DefaultColWidthRecord rec = new DefaultColWidthRecord();
            rec.field_1_col_width = field_1_col_width;
            return rec;
        }
    }
}
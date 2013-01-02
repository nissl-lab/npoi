
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
     * Title:        GridSet Record.
     * Description:  flag denoting whether the user specified that gridlines are used when
     *               printing.
     * REFERENCE:  PG 320 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     *
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author  Glen Stampoultzis (glens at apache.org)
     * @author Jason Height (jheight at chariot dot net dot au)
     *
     * @version 2.0-pre
     */

    public class GridsetRecord
       : StandardRecord
    {
        public const short sid = 0x82;
        public short field_1_gridset_flag;

        public GridsetRecord()
        {
        }

        /**
         * Constructs a GridSet record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public GridsetRecord(RecordInputStream in1)
        {
            field_1_gridset_flag = in1.ReadShort();
        }

        /**
         * Get whether the gridlines are shown during printing.
         *
         * @return gridSet - true if gridlines are NOT printed, false if they are.
         */

        public bool Gridset
        {
            get
            {
                return (field_1_gridset_flag == 1);
            }
            set
            {
                if (value == true)
                {
                    field_1_gridset_flag = 1;
                }
                else
                {
                    field_1_gridset_flag = 0;
                }
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[GRIDSET]\n");
            buffer.Append("    .gridset        = ").Append(Gridset)
                .Append("\n");
            buffer.Append("[/GRIDSET]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_gridset_flag);
        }

        protected override int DataSize
        {
            get
            {
                return 2;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            GridsetRecord rec = new GridsetRecord();
            rec.field_1_gridset_flag = field_1_gridset_flag;
            return rec;
        }
    }
}

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

namespace NPOI.HSSF.Record.Chart
{
    using System.Collections;
    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * The number format index record indexes format table.  This applies to an axis.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class NumberFormatIndexRecord
       : StandardRecord
    {
        public const short sid = 0x104e;
        private short field_1_formatIndex;


        public NumberFormatIndexRecord()
        {

        }

        /**
         * Constructs a NumberFormatIndex record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public NumberFormatIndexRecord(RecordInputStream in1)
        {
            field_1_formatIndex = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[IFMT]\n");
            buffer.Append("    .formatIndex          = ")
                .Append("0x").Append(HexDump.ToHex(FormatIndex))
                .Append(" (").Append(FormatIndex).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/IFMT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_formatIndex);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
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
            NumberFormatIndexRecord rec = new NumberFormatIndexRecord();

            rec.field_1_formatIndex = field_1_formatIndex;
            return rec;
        }




        /**
         * Get the format index field for the NumberFormatIndex record.
         */
        public short FormatIndex
        {
            get { return field_1_formatIndex; }
            set { this.field_1_formatIndex = value; }
        }


    }
}




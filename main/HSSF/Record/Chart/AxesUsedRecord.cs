
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

    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * The number of axes used on a chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class AxesUsedRecord
       : StandardRecord
    {
        public const short sid = 0x1046;
        private short field_1_numAxis;


        public AxesUsedRecord()
        {

        }

        /**
         * Constructs a AxisUsed record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public AxesUsedRecord(RecordInputStream in1)
        {
            field_1_numAxis = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AXISUSED]\n");
            buffer.Append("    .numAxis              = ")
                .Append("0x").Append(HexDump.ToHex(this.NumAxis))
                .Append(" (").Append(this.NumAxis).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/AXISUSED]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_numAxis);
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
            AxesUsedRecord rec = new AxesUsedRecord();

            rec.field_1_numAxis = field_1_numAxis;
            return rec;
        }




        /**
         * Get the num axis field for the AxisUsed record.
         */
        public short NumAxis
        {
            get { return field_1_numAxis; }
            set { this.field_1_numAxis = value; }
        }


    }
}





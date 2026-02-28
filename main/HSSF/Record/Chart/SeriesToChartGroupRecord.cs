
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
     * Indicates the chart-group index for a series.  The order probably defines the mapping.  So the 0th record probably means the 0th series.  The only field in this of course defines which chart Group the 0th series (for instance) would map to.  Confusing?  Well thats because it Is.  (p 522 BCG)
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    //
    [Obsolete("duplication record, see SerToCrtRecord")]
    public class SeriesToChartGroupRecord
       : StandardRecord
    {
        public const short sid = 0x1045;
        private short field_1_chartGroupIndex;


        public SeriesToChartGroupRecord()
        {

        }

        /**
         * Constructs a SeriesToChartGroup record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SeriesToChartGroupRecord(RecordInputStream in1)
        {

            field_1_chartGroupIndex = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SeriesToChartGroup]\n");
            buffer.Append("    .chartGroupIndex      = ")
                .Append("0x").Append(HexDump.ToHex(ChartGroupIndex))
                .Append(" (").Append(ChartGroupIndex).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SeriesToChartGroup]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_chartGroupIndex);
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
            SeriesToChartGroupRecord rec = new SeriesToChartGroupRecord();

            rec.field_1_chartGroupIndex = field_1_chartGroupIndex;
            return rec;
        }




        /**
         * Get the chart Group index field for the SeriesToChartGroup record.
         */
        public short ChartGroupIndex
        {
            get { return field_1_chartGroupIndex; }
            set { this.field_1_chartGroupIndex = value; }
        }


    }
}

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
     * The area record is used to define a area chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class AreaRecord
       : StandardRecord
    {
        public const short sid = 0x101A;
        private short field_1_formatFlags;
        private BitField stacked = BitFieldFactory.GetInstance(0x1);
        private BitField DisplayAsPercentage = BitFieldFactory.GetInstance(0x2);
        private BitField shadow = BitFieldFactory.GetInstance(0x4);


        public AreaRecord()
        {

        }

        /**
         * Constructs a Area record and s its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public AreaRecord(RecordInputStream in1)
        {
               field_1_formatFlags = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AREA]\n");
            buffer.Append("    .formatFlags          = ")
                .Append("0x").Append(HexDump.ToHex(FormatFlags))
                .Append(" (").Append(FormatFlags).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .stacked                  = ").Append(IsStacked).Append('\n');
            buffer.Append("         .DisplayAsPercentage      = ").Append(IsDisplayAsPercentage).Append('\n');
            buffer.Append("         .shadow                   = ").Append(IsShadow).Append('\n');

            buffer.Append("[/AREA]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_formatFlags);
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
            AreaRecord rec = new AreaRecord();

            rec.field_1_formatFlags = field_1_formatFlags;
            return rec;
        }




        /**
         *  the format flags field for the Area record.
         */
        public short FormatFlags
        {
            get { return field_1_formatFlags; }
            set { this.field_1_formatFlags = value; }
        }

        /**
         * series is stacked
         * @return  the stacked field value.
         */
        public bool IsStacked
        {
            get { return stacked.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = stacked.SetShortBoolean(field_1_formatFlags, value); }
        }


        /**
         * results Displayed as percentages
         * @return  the Display as percentage field value.
         */
        public bool IsDisplayAsPercentage
        {
            get { return DisplayAsPercentage.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = DisplayAsPercentage.SetShortBoolean(field_1_formatFlags, value); }
        }

        /**
         * Display a shadow for the chart
         * @return  the shadow field value.
         */
        public bool IsShadow
        {
            get { return shadow.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = shadow.SetShortBoolean(field_1_formatFlags, value); }
        }


    }
}





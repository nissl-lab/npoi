
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
     * The bar record is used to define a bar chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class BarRecord
       : StandardRecord
    {
        public const short sid = 0x1017;
        private short field_1_barSpace;
        private short field_2_categorySpace;
        private short field_3_formatFlags;
        private BitField horizontal = BitFieldFactory.GetInstance(0x1);
        private BitField stacked = BitFieldFactory.GetInstance(0x2);
        private BitField DisplayAsPercentage = BitFieldFactory.GetInstance(0x4);
        private BitField shadow = BitFieldFactory.GetInstance(0x8);


        public BarRecord()
        {

        }

        /**
         * Constructs a Bar record and s its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public BarRecord(RecordInputStream in1)
        {

            field_1_barSpace = in1.ReadShort();
            field_2_categorySpace = in1.ReadShort();
            field_3_formatFlags = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[BAR]\n");
            buffer.Append("    .barSpace             = ")
                .Append("0x").Append(HexDump.ToHex(BarSpace))
                .Append(" (").Append(BarSpace).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .categorySpace        = ")
                .Append("0x").Append(HexDump.ToHex(CategorySpace))
                .Append(" (").Append(CategorySpace).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .formatFlags          = ")
                .Append("0x").Append(HexDump.ToHex(FormatFlags))
                .Append(" (").Append(FormatFlags).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .horizontal               = ").Append(IsHorizontal).Append('\n');
            buffer.Append("         .stacked                  = ").Append(IsStacked).Append('\n');
            buffer.Append("         .DisplayAsPercentage      = ").Append(IsDisplayAsPercentage).Append('\n');
            buffer.Append("         .shadow                   = ").Append(IsShadow).Append('\n');

            buffer.Append("[/BAR]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_barSpace);
            out1.WriteShort(field_2_categorySpace);
            out1.WriteShort(field_3_formatFlags);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            BarRecord rec = new BarRecord();

            rec.field_1_barSpace = field_1_barSpace;
            rec.field_2_categorySpace = field_2_categorySpace;
            rec.field_3_formatFlags = field_3_formatFlags;
            return rec;
        }




        /**
         *  the bar space field for the Bar record.
         */
        public short BarSpace
        {
            get { return field_1_barSpace; }
            set { this.field_1_barSpace = value; }
        }

        /**
         *  the category space field for the Bar record.
         */
        public short CategorySpace
        {
            get { return field_2_categorySpace; }
            set { this.field_2_categorySpace = value; }
        }

        /**
         *  the format flags field for the Bar record.
         */
        public short FormatFlags
        {
            get { return field_3_formatFlags; }
            set { this.field_3_formatFlags = value; }
        }

        /**
         * true to Display horizontal bar charts, false for vertical
         * @return  the horizontal field value.
         */
        public bool IsHorizontal
        {
            get { return horizontal.IsSet(field_3_formatFlags); }
            set { field_3_formatFlags = horizontal.SetShortBoolean(field_3_formatFlags, value); }
        }

        /**
         * stack Displayed values
         * @return  the stacked field value.
         */
        public bool IsStacked
        {
            get { return stacked.IsSet(field_3_formatFlags); }
            set { field_3_formatFlags = stacked.SetShortBoolean(field_3_formatFlags, value); }
        }


        /**
         * Display chart values as a percentage
         * @return  the Display as percentage field value.
         */
        public bool IsDisplayAsPercentage
        {
            get { return DisplayAsPercentage.IsSet(field_3_formatFlags); }
            set { field_3_formatFlags = DisplayAsPercentage.SetShortBoolean(field_3_formatFlags, value); }
        }

        /**
         * Display a shadow for the chart
         * @return  the shadow field value.
         */
        public bool IsShadow
        {
            get { return shadow.IsSet(field_3_formatFlags); }
            set { field_3_formatFlags = shadow.SetShortBoolean(field_3_formatFlags, value); }
        }


    }
}





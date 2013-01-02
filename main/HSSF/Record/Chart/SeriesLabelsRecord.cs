
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
     * The series label record defines the type of label associated with the data format record.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class SeriesLabelsRecord
       : StandardRecord
    {
        public static short sid = 0x100c;
        private short field_1_formatFlags;
        private BitField showActual = BitFieldFactory.GetInstance(0x1);
        private BitField showPercent = BitFieldFactory.GetInstance(0x2);
        private BitField labelAsPercentage = BitFieldFactory.GetInstance(0x4);
        private BitField smoothedLine = BitFieldFactory.GetInstance(0x8);
        private BitField showLabel = BitFieldFactory.GetInstance(0x10);
        private BitField showBubbleSizes = BitFieldFactory.GetInstance(0x20);


        public SeriesLabelsRecord()
        {

        }

        /**
         * Constructs a SeriesLabels record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SeriesLabelsRecord(RecordInputStream in1)
        {
            field_1_formatFlags = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ATTACHEDLABEL]\n");
            buffer.Append("    .formatFlags          = ")
                .Append("0x").Append(HexDump.ToHex(FormatFlags))
                .Append(" (").Append(FormatFlags).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .showActual               = ").Append(IsShowActual).Append('\n');
            buffer.Append("         .showPercent              = ").Append(IsShowPercent).Append('\n');
            buffer.Append("         .labelAsPercentage        = ").Append(IsLabelAsPercentage).Append('\n');
            buffer.Append("         .smoothedLine             = ").Append(IsSmoothedLine).Append('\n');
            buffer.Append("         .showLabel                = ").Append(IsShowLabel).Append('\n');
            buffer.Append("         .showBubbleSizes          = ").Append(IsShowBubbleSizes).Append('\n');

            buffer.Append("[/ATTACHEDLABEL]\n");
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
            SeriesLabelsRecord rec = new SeriesLabelsRecord();

            rec.field_1_formatFlags = field_1_formatFlags;
            return rec;
        }




        /**
         * Get the format flags field for the SeriesLabels record.
         */
        public short FormatFlags
        {
            get { return field_1_formatFlags; }
            set { this.field_1_formatFlags = value; }
        }

        /**
         * show actual value of the data point
         * @return  the show actual field value.
         */
        public bool IsShowActual
        {
            get { return showActual.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = showActual.SetShortBoolean(field_1_formatFlags, value); }
        }

        /**
         * show value as percentage of total (pie charts only)
         * @return  the show percent field value.
         */
        public bool IsShowPercent
        {
            get { return showPercent.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = showPercent.SetShortBoolean(field_1_formatFlags, value); }
        }

        /**
         * show category label/value as percentage (pie charts only)
         * @return  the label as percentage field value.
         */
        public bool IsLabelAsPercentage
        {
            get { return labelAsPercentage.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = labelAsPercentage.SetShortBoolean(field_1_formatFlags, value); }
        }

        /**
         * show smooth line
         * @return  the smoothed line field value.
         */
        public bool IsSmoothedLine
        {
            get { return smoothedLine.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = smoothedLine.SetShortBoolean(field_1_formatFlags, value); }
        }
        /**
         * Display category label
         * @return  the show label field value.
         */
        public bool IsShowLabel
        {
            get { return showLabel.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = showLabel.SetShortBoolean(field_1_formatFlags, value); }
        }

        /**
         * ??
         * @return  the show bubble sizes field value.
         */
        public bool IsShowBubbleSizes
        {
            get { return showBubbleSizes.IsSet(field_1_formatFlags); }
            set { field_1_formatFlags = showBubbleSizes.SetShortBoolean(field_1_formatFlags, value); }
        }


    }
}

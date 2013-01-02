
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

    
    using System.Text;
    using System;
    using NPOI.Util;


    /**
     * The dat record is used to store options for the chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class DatRecord
       : StandardRecord
    {
        public const short sid = 0x1063;
        private short field_1_options;
        private BitField horizontalBorder = BitFieldFactory.GetInstance(0x1);
        private BitField verticalBorder = BitFieldFactory.GetInstance(0x2);
        private BitField border = BitFieldFactory.GetInstance(0x4);
        private BitField showSeriesKey = BitFieldFactory.GetInstance(0x8);


        public DatRecord()
        {

        }

        /**
         * Constructs a Dat record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public DatRecord(RecordInputStream in1)
        {
            field_1_options = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DAT]\n");
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .horizontalBorder         = ").Append(IsHorizontalBorder()).Append('\n');
            buffer.Append("         .verticalBorder           = ").Append(IsVerticalBorder()).Append('\n');
            buffer.Append("         .border                   = ").Append(IsBorder()).Append('\n');
            buffer.Append("         .showSeriesKey            = ").Append(IsShowSeriesKey()).Append('\n');

            buffer.Append("[/DAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_options);
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
            DatRecord rec = new DatRecord();

            rec.field_1_options = field_1_options;
            return rec;
        }




        /**
         * Get the options field for the Dat record.
         */
        public short Options
        {
            get { return field_1_options; }
            set { this.field_1_options = value; }
        }

        /**
         * Sets the horizontal border field value.
         * has a horizontal border
         */
        public void SetHorizontalBorder(bool value)
        {
            field_1_options = horizontalBorder.SetShortBoolean(field_1_options, value);
        }

        /**
         * has a horizontal border
         * @return  the horizontal border field value.
         */
        public bool IsHorizontalBorder()
        {
            return horizontalBorder.IsSet(field_1_options);
        }

        /**
         * Sets the vertical border field value.
         * has vertical border
         */
        public void SetVerticalBorder(bool value)
        {
            field_1_options = verticalBorder.SetShortBoolean(field_1_options, value);
        }

        /**
         * has vertical border
         * @return  the vertical border field value.
         */
        public bool IsVerticalBorder()
        {
            return verticalBorder.IsSet(field_1_options);
        }

        /**
         * Sets the border field value.
         * data table has a border
         */
        public void SetBorder(bool value)
        {
            field_1_options = border.SetShortBoolean(field_1_options, value);
        }

        /**
         * data table has a border
         * @return  the border field value.
         */
        public bool IsBorder()
        {
            return border.IsSet(field_1_options);
        }

        /**
         * Sets the show series key field value.
         * shows the series key
         */
        public void SetShowSeriesKey(bool value)
        {
            field_1_options = showSeriesKey.SetShortBoolean(field_1_options, value);
        }

        /**
         * shows the series key
         * @return  the show series key field value.
         */
        public bool IsShowSeriesKey()
        {
            return showSeriesKey.IsSet(field_1_options);
        }


    }
}




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
     * The data format record is used to index into a series.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class DataFormatRecord
       : StandardRecord
    {
        public const short sid = 0x1006;
        private short field_1_pointNumber;
        private short field_2_seriesIndex;
        private short field_3_seriesNumber;
        private short field_4_formatFlags;
        private BitField useExcel4Colors = BitFieldFactory.GetInstance(0x1);


        public DataFormatRecord()
        {

        }

        /**
         * Constructs a DataFormat record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public DataFormatRecord(RecordInputStream in1)
        {
            field_1_pointNumber = in1.ReadShort();
            field_2_seriesIndex = in1.ReadShort();
            field_3_seriesNumber = in1.ReadShort();
            field_4_formatFlags = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DATAFORMAT]\n");
            buffer.Append("    .pointNumber          = ")
                .Append("0x").Append(HexDump.ToHex(PointNumber))
                .Append(" (").Append(PointNumber).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .seriesIndex          = ")
                .Append("0x").Append(HexDump.ToHex(SeriesIndex))
                .Append(" (").Append(SeriesIndex).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .seriesNumber         = ")
                .Append("0x").Append(HexDump.ToHex(SeriesNumber))
                .Append(" (").Append(SeriesNumber).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .formatFlags          = ")
                .Append("0x").Append(HexDump.ToHex(FormatFlags))
                .Append(" (").Append(FormatFlags).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .useExcel4Colors          = ").Append(UseExcel4Colors).Append('\n');

            buffer.Append("[/DATAFORMAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_pointNumber);
            out1.WriteShort(field_2_seriesIndex);
            out1.WriteShort(field_3_seriesNumber);
            out1.WriteShort(field_4_formatFlags);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return  2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            DataFormatRecord rec = new DataFormatRecord();

            rec.field_1_pointNumber = field_1_pointNumber;
            rec.field_2_seriesIndex = field_2_seriesIndex;
            rec.field_3_seriesNumber = field_3_seriesNumber;
            rec.field_4_formatFlags = field_4_formatFlags;
            return rec;
        }




        /**
         * Get the point number field for the DataFormat record.
         */
        public short PointNumber
        {
            get
            {
                return field_1_pointNumber;
            }
            set { this.field_1_pointNumber = value; }
        }

        /**
         * Get the series index field for the DataFormat record.
         */
        public short SeriesIndex
        {
            get
            {
                return field_2_seriesIndex;
            }
            set { this.field_2_seriesIndex = value; }
        }


        /**
         * Get the series number field for the DataFormat record.
         */
        public short SeriesNumber
        {
            get
            {
                return field_3_seriesNumber;
            }
            set { this.field_3_seriesNumber = value; }
        }

        /**
         * Get the format flags field for the DataFormat record.
         */
        public short FormatFlags
        {
            get { return field_4_formatFlags; }
            set { this.field_4_formatFlags = value; }
        }

        /**
         * Set true to use excel 4 colors.
         * @return  the use excel 4 colors field value.
         */
        public bool UseExcel4Colors
        {
            get
            {
                return useExcel4Colors.IsSet(field_4_formatFlags);
            }
            set { field_4_formatFlags = useExcel4Colors.SetShortBoolean(field_4_formatFlags, value); }
        }
    }
}

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
     * The series record describes the overall data for a series.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class SeriesRecord
       : StandardRecord
    {
        public const short sid = 0x1003;
        private short field_1_categoryDataType;
        public const short CATEGORY_DATA_TYPE_DATES = 0;
        public const short CATEGORY_DATA_TYPE_NUMERIC = 1;
        public const short CATEGORY_DATA_TYPE_SEQUENCE = 2;
        public const short CATEGORY_DATA_TYPE_TEXT = 3;
        private short field_2_valuesDataType;
        public const short VALUES_DATA_TYPE_DATES = 0;
        public const short VALUES_DATA_TYPE_NUMERIC = 1;
        public const short VALUES_DATA_TYPE_SEQUENCE = 2;
        public const short VALUES_DATA_TYPE_TEXT = 3;
        private short field_3_numCategories;
        private short field_4_numValues;
        private short field_5_bubbleSeriesType;
        public const short BUBBLE_SERIES_TYPE_DATES = 0;
        public const short BUBBLE_SERIES_TYPE_NUMERIC = 1;
        public const short BUBBLE_SERIES_TYPE_SEQUENCE = 2;
        public const short BUBBLE_SERIES_TYPE_TEXT = 3;
        private short field_6_numBubbleValues;


        public SeriesRecord()
        {

        }

        /**
         * Constructs a Series record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SeriesRecord(RecordInputStream in1)
        {
            field_1_categoryDataType = in1.ReadShort();
            field_2_valuesDataType = in1.ReadShort();
            field_3_numCategories = in1.ReadShort();
            field_4_numValues = in1.ReadShort();
            field_5_bubbleSeriesType = in1.ReadShort();
            field_6_numBubbleValues = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SERIES]\n");
            buffer.Append("    .categoryDataType     = ")
                .Append("0x").Append(HexDump.ToHex(CategoryDataType))
                .Append(" (").Append(CategoryDataType).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .valuesDataType       = ")
                .Append("0x").Append(HexDump.ToHex(ValuesDataType))
                .Append(" (").Append(ValuesDataType).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .numCategories        = ")
                .Append("0x").Append(HexDump.ToHex(NumCategories))
                .Append(" (").Append(NumCategories).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .numValues            = ")
                .Append("0x").Append(HexDump.ToHex(NumValues))
                .Append(" (").Append(NumValues).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .bubbleSeriesType     = ")
                .Append("0x").Append(HexDump.ToHex(BubbleSeriesType))
                .Append(" (").Append(BubbleSeriesType).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .numBubbleValues      = ")
                .Append("0x").Append(HexDump.ToHex(NumBubbleValues))
                .Append(" (").Append(NumBubbleValues).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SERIES]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_categoryDataType);
            out1.WriteShort(field_2_valuesDataType);
            out1.WriteShort(field_3_numCategories);
            out1.WriteShort(field_4_numValues);
            out1.WriteShort(field_5_bubbleSeriesType);
            out1.WriteShort(field_6_numBubbleValues);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2 + 2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            SeriesRecord rec = new SeriesRecord();

            rec.field_1_categoryDataType = field_1_categoryDataType;
            rec.field_2_valuesDataType = field_2_valuesDataType;
            rec.field_3_numCategories = field_3_numCategories;
            rec.field_4_numValues = field_4_numValues;
            rec.field_5_bubbleSeriesType = field_5_bubbleSeriesType;
            rec.field_6_numBubbleValues = field_6_numBubbleValues;
            return rec;
        }




        /**
         * Get the category data type field for the Series record.
         *
         * @return  One of 
         *        CATEGORY_DATA_TYPE_DATES
         *        CATEGORY_DATA_TYPE_NUMERIC
         *        CATEGORY_DATA_TYPE_SEQUENCE
         *        CATEGORY_DATA_TYPE_TEXT
         */
        public short CategoryDataType
        {
            get
            {
                return field_1_categoryDataType;
            }
            set 
            {
                this.field_1_categoryDataType = value;
            }
        }
        /**
         * Get the values data type field for the Series record.
         *
         * @return  One of 
         *        VALUES_DATA_TYPE_DATES
         *        VALUES_DATA_TYPE_NUMERIC
         *        VALUES_DATA_TYPE_SEQUENCE
         *        VALUES_DATA_TYPE_TEXT
         */
        public short ValuesDataType
        {
            get
            {
                return field_2_valuesDataType;
            }
            set 
            {
                this.field_2_valuesDataType = value;
            }
        }


        /**
         * Get the num categories field for the Series record.
         */
        public short NumCategories
        {
            get
            {
                return field_3_numCategories;
            }
            set 
            {
                this.field_3_numCategories = value;
            }
        }

        /**
         * Get the num values field for the Series record.
         */
        public short NumValues
        {
            get
            {
                return field_4_numValues;
            }
            set 
            {
                this.field_4_numValues = value;
            }
        }

        /**
         * Get the bubble series type field for the Series record.
         *
         * @return  One of 
         *        BUBBLE_SERIES_TYPE_DATES
         *        BUBBLE_SERIES_TYPE_NUMERIC
         *        BUBBLE_SERIES_TYPE_SEQUENCE
         *        BUBBLE_SERIES_TYPE_TEXT
         */
        public short BubbleSeriesType
        {
            get
            {
                return field_5_bubbleSeriesType;
            }
            set 
            {
                this.field_5_bubbleSeriesType = value;
            }
        }
        /**
         * Get the num bubble values field for the Series record.
         */
        public short NumBubbleValues
        {
            get
            {
                return field_6_numBubbleValues;
            }
            set 
            {
                this.field_6_numBubbleValues =value;
            }
        }


    }
}
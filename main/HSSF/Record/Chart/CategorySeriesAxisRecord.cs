
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
    using System.IO;
    using System.Text;
    using NPOI.Util;

    /**
     * This record refers to a category or series axis and is used to specify label/tickmark frequency.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class CategorySeriesAxisRecord
       : StandardRecord
    {
        public const short sid = 0x1020;
        private short field_1_crossingPoint;
        private short field_2_labelFrequency;
        private short field_3_tickMarkFrequency;
        private short field_4_options;
        private BitField valueAxisCrossing = BitFieldFactory.GetInstance(0x1);
        private BitField crossesFarRight = BitFieldFactory.GetInstance(0x2);
        private BitField reversed = BitFieldFactory.GetInstance(0x4);


        public CategorySeriesAxisRecord()
        {

        }

        /**
         * Constructs a CategorySeriesAxis record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public CategorySeriesAxisRecord(RecordInputStream in1)
    {
        field_1_crossingPoint = in1.ReadShort();
        field_2_labelFrequency = in1.ReadShort();
        field_3_tickMarkFrequency = in1.ReadShort();
        field_4_options = in1.ReadShort(); 
    
    }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CATSERRANGE]\n");
            buffer.Append("    .crossingPoint        = ")
                .Append("0x").Append(HexDump.ToHex(CrossingPoint))
                .Append(" (").Append(CrossingPoint).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .labelFrequency       = ")
                .Append("0x").Append(HexDump.ToHex(LabelFrequency))
                .Append(" (").Append(LabelFrequency).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .tickMarkFrequency    = ")
                .Append("0x").Append(HexDump.ToHex(TickMarkFrequency))
                .Append(" (").Append(TickMarkFrequency).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .valueAxisCrossing        = ").Append(IsValueAxisCrossing).Append('\n');
            buffer.Append("         .crossesFarRight          = ").Append(IsCrossesFarRight).Append('\n');
            buffer.Append("         .reversed                 = ").Append(IsReversed).Append('\n');

            buffer.Append("[/CATSERRANGE]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_crossingPoint);
            out1.WriteShort(field_2_labelFrequency);
            out1.WriteShort(field_3_tickMarkFrequency);
            out1.WriteShort(field_4_options);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            CategorySeriesAxisRecord rec = new CategorySeriesAxisRecord();

            rec.field_1_crossingPoint = field_1_crossingPoint;
            rec.field_2_labelFrequency = field_2_labelFrequency;
            rec.field_3_tickMarkFrequency = field_3_tickMarkFrequency;
            rec.field_4_options = field_4_options;
            return rec;
        }




        /**
         * Get the crossing point field for the CategorySeriesAxis record.
         */
        public short CrossingPoint
        {
            get
            {
                return field_1_crossingPoint;
            }
            set 
            {
                this.field_1_crossingPoint = value;
            }
        }

        /**
         * Get the label frequency field for the CategorySeriesAxis record.
         */
        public short LabelFrequency
        {
            get
            {
                return field_2_labelFrequency;
            }
            set 
            {
                this.field_2_labelFrequency = value;
            }
        }

        /**
         * Get the tick mark frequency field for the CategorySeriesAxis record.
         */
        public short TickMarkFrequency
        {
            get
            {
                return field_3_tickMarkFrequency;
            }
            set 
            {
                this.field_3_tickMarkFrequency = value;
            }
        }


        /**
         * Get the options field for the CategorySeriesAxis record.
         */
        public short Options
        {
            get { return field_4_options; }
            set { this.field_4_options = value; }
        }

        /**
         * Set true to indicate axis crosses between categories and false to cross axis midway
         * @return  the value axis crossing field value.
         */
        public bool IsValueAxisCrossing
        {
            get
            {
                return valueAxisCrossing.IsSet(field_4_options);
            }
            set 
            {
                field_4_options = valueAxisCrossing.SetShortBoolean(field_4_options, value);
            }
        }

        /**
         * axis crosses at the far right
         * @return  the crosses far right field value.
         */
        public bool IsCrossesFarRight
        {
            get
            {
                return crossesFarRight.IsSet(field_4_options);
            }
            set 
            {
                field_4_options = crossesFarRight.SetShortBoolean(field_4_options, value);
            }
        }

        /**
         * categories are Displayed in reverse order
         * @return  the reversed field value.
         */
        public bool IsReversed
        {
            get
            {
                return reversed.IsSet(field_4_options);
            }
            set 
            {
                field_4_options = reversed.SetShortBoolean(field_4_options, value);
            }
        }


    }
}



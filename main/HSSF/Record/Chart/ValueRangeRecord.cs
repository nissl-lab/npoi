
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
     * The value range record defines the range of the value axis.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class ValueRangeRecord
       : StandardRecord
    {
        public const short sid = 0x101f;
        private double field_1_minimumAxisValue;
        private double field_2_maximumAxisValue;
        private double field_3_majorIncrement;
        private double field_4_minorIncrement;
        private double field_5_categoryAxisCross;
        private short field_6_options;
        private BitField automaticMinimum = BitFieldFactory.GetInstance(0x1);
        private BitField automaticMaximum = BitFieldFactory.GetInstance(0x2);
        private BitField automaticMajor = BitFieldFactory.GetInstance(0x4);
        private BitField automaticMinor = BitFieldFactory.GetInstance(0x8);
        private BitField automaticCategoryCrossing = BitFieldFactory.GetInstance(0x10);
        private BitField logarithmicScale = BitFieldFactory.GetInstance(0x20);
        private BitField valuesInReverse = BitFieldFactory.GetInstance(0x40);
        private BitField crossCategoryAxisAtMaximum = BitFieldFactory.GetInstance(0x80);
        private BitField reserved = BitFieldFactory.GetInstance(0x100);


        public ValueRangeRecord()
        {

        }

        /**
         * Constructs a ValueRange record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public ValueRangeRecord(RecordInputStream in1)
        {
            field_1_minimumAxisValue = in1.ReadDouble();
            field_2_maximumAxisValue = in1.ReadDouble();
            field_3_majorIncrement = in1.ReadDouble();
            field_4_minorIncrement = in1.ReadDouble();
            field_5_categoryAxisCross = in1.ReadDouble();
            field_6_options = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[VALUERANGE]\n");
            buffer.Append("    .minimumAxisValue     = ")
                .Append(" (").Append(MinimumAxisValue).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .maximumAxisValue     = ")
                .Append(" (").Append(MaximumAxisValue).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .majorIncrement       = ")
                .Append(" (").Append(MajorIncrement).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .minorIncrement       = ")
                .Append(" (").Append(MinorIncrement).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .categoryAxisCross    = ")
                .Append(" (").Append(CategoryAxisCross).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .automaticMinimum         = ").Append(IsAutomaticMinimum).Append('\n');
            buffer.Append("         .automaticMaximum         = ").Append(IsAutomaticMaximum).Append('\n');
            buffer.Append("         .automaticMajor           = ").Append(IsAutomaticMajor).Append('\n');
            buffer.Append("         .automaticMinor           = ").Append(IsAutomaticMinor).Append('\n');
            buffer.Append("         .automaticCategoryCrossing     = ").Append(IsAutomaticCategoryCrossing).Append('\n');
            buffer.Append("         .logarithmicScale         = ").Append(IsLogarithmicScale).Append('\n');
            buffer.Append("         .valuesInReverse          = ").Append(IsValuesInReverse).Append('\n');
            buffer.Append("         .crossCategoryAxisAtMaximum     = ").Append(IsCrossCategoryAxisAtMaximum).Append('\n');
            buffer.Append("         .reserved                 = ").Append(IsReserved).Append('\n');

            buffer.Append("[/VALUERANGE]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteDouble(field_1_minimumAxisValue);
            out1.WriteDouble(field_2_maximumAxisValue);
            out1.WriteDouble(field_3_majorIncrement);
            out1.WriteDouble(field_4_minorIncrement);
            out1.WriteDouble(field_5_categoryAxisCross);
            out1.WriteShort(field_6_options);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 8 + 8 + 8 + 8 + 8 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            ValueRangeRecord rec = new ValueRangeRecord();

            rec.field_1_minimumAxisValue = field_1_minimumAxisValue;
            rec.field_2_maximumAxisValue = field_2_maximumAxisValue;
            rec.field_3_majorIncrement = field_3_majorIncrement;
            rec.field_4_minorIncrement = field_4_minorIncrement;
            rec.field_5_categoryAxisCross = field_5_categoryAxisCross;
            rec.field_6_options = field_6_options;
            return rec;
        }




        /**
         * Get the minimum axis value field for the ValueRange record.
         */
        public double MinimumAxisValue
        {
            get{return field_1_minimumAxisValue;}
            set { this.field_1_minimumAxisValue = value; }
        }

        /**
         * Get the maximum axis value field for the ValueRange record.
         */
        public double MaximumAxisValue
        {
            get{return field_2_maximumAxisValue;}
            set { this.field_2_maximumAxisValue = value; }
        }


        /**
         * Get the major increment field for the ValueRange record.
         */
        public double MajorIncrement
        {
            get { return field_3_majorIncrement; }
            set { this.field_3_majorIncrement = value; }
        }

        /**
         * Get the minor increment field for the ValueRange record.
         */
        public double MinorIncrement
        {
            get{return field_4_minorIncrement;}
            set { this.field_4_minorIncrement = value; }
        }


        /**
         * Get the category axis cross field for the ValueRange record.
         */
        public double CategoryAxisCross
        {
            get
            {
                return field_5_categoryAxisCross;
            }
            set { this.field_5_categoryAxisCross = value; }
        }

        /**
         * Get the options field for the ValueRange record.
         */
        public short Options
        {
            get{return field_6_options;}
            set { this.field_6_options = value; }
        }


        /**
         * automatic minimum value selected
         * @return  the automatic minimum field value.
         */
        public bool IsAutomaticMinimum
        {
            get
            {
                return automaticMinimum.IsSet(field_6_options);
            }
            set { field_6_options = automaticMinimum.SetShortBoolean(field_6_options, value); }
        }
        /**
         * automatic maximum value selected
         * @return  the automatic maximum field value.
         */
        public bool IsAutomaticMaximum
        {
            get { return automaticMaximum.IsSet(field_6_options); }
            set { field_6_options = automaticMaximum.SetShortBoolean(field_6_options, value); }
        }

        /**
         * automatic major Unit selected
         * @return  the automatic major field value.
         */
        public bool IsAutomaticMajor
        {
            get { return automaticMajor.IsSet(field_6_options); }
            set { field_6_options = automaticMajor.SetShortBoolean(field_6_options, value); }
        }

        /**
         * automatic minor Unit selected
         * @return  the automatic minor field value.
         */
        public bool IsAutomaticMinor
        {
            get{return automaticMinor.IsSet(field_6_options);}
            set { field_6_options = automaticMinor.SetShortBoolean(field_6_options, value); }
        }

        /**
         * category crossing point is automatically selected
         * @return  the automatic category crossing field value.
         */
        public bool IsAutomaticCategoryCrossing
        {
            get { return automaticCategoryCrossing.IsSet(field_6_options); }
            set { field_6_options = automaticCategoryCrossing.SetShortBoolean(field_6_options, value); }
        }

        /**
         * use logarithmic scale
         * @return  the logarithmic scale field value.
         */
        public bool IsLogarithmicScale
        {
            get{return logarithmicScale.IsSet(field_6_options);}
            set { field_6_options = logarithmicScale.SetShortBoolean(field_6_options, value); }
        }

        /**
         * values are reverses in graph
         * @return  the values in reverse field value.
         */
        public bool IsValuesInReverse
        {
            get { return valuesInReverse.IsSet(field_6_options); }
            set { field_6_options = valuesInReverse.SetShortBoolean(field_6_options, value); }
        }

        /**
         * category axis to cross at maximum value
         * @return  the cross category axis at maximum field value.
         */
        public bool IsCrossCategoryAxisAtMaximum
        {
            get
            {
                return crossCategoryAxisAtMaximum.IsSet(field_6_options);
            }
            set { field_6_options = crossCategoryAxisAtMaximum.SetShortBoolean(field_6_options, value); }
        }


        /**
         * reserved, must equal 1 (excel dev. guide says otherwise)
         * @return  the reserved field value.
         */
        public bool IsReserved
        {
            get{return reserved.IsSet(field_6_options);}
            set { field_6_options = reserved.SetShortBoolean(field_6_options, value); }
        }


    }
}
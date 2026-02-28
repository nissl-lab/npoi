/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record.Chart
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;


    /**
     * The axis options record provides unit information and other various tidbits about the axis.<p/>
     * 
     * @author Andrew C. Oliver(acoliver at apache.org)
     */
    public class AxisOptionsRecord : StandardRecord, ICloneable
    {
        public static short sid = 0x1062;

        private static BitField defaultMinimum = BitFieldFactory.GetInstance(0x01);
        private static BitField defaultMaximum = BitFieldFactory.GetInstance(0x02);
        private static BitField defaultMajor = BitFieldFactory.GetInstance(0x04);
        private static BitField defaultMinorUnit = BitFieldFactory.GetInstance(0x08);
        private static BitField isDate = BitFieldFactory.GetInstance(0x10);
        private static BitField defaultBase = BitFieldFactory.GetInstance(0x20);
        private static BitField defaultCross = BitFieldFactory.GetInstance(0x40);
        private static BitField defaultDateSettings = BitFieldFactory.GetInstance(0x80);

        private short field_1_minimumCategory;
        private short field_2_maximumCategory;
        private short field_3_majorUnitValue;
        private short field_4_majorUnit;
        private short field_5_minorUnitValue;
        private short field_6_minorUnit;
        private short field_7_baseUnit;
        private short field_8_crossingPoint;
        private short field_9_options;


        public AxisOptionsRecord()
        {

        }

        public AxisOptionsRecord(RecordInputStream in1)
        {
            field_1_minimumCategory = in1.ReadShort();
            field_2_maximumCategory = in1.ReadShort();
            field_3_majorUnitValue = in1.ReadShort();
            field_4_majorUnit = in1.ReadShort();
            field_5_minorUnitValue = in1.ReadShort();
            field_6_minorUnit = in1.ReadShort();
            field_7_baseUnit = in1.ReadShort();
            field_8_crossingPoint = in1.ReadShort();
            field_9_options = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AXCEXT]\n");
            buffer.Append("    .minimumCategory      = ")
                .Append("0x").Append(HexDump.ToHex(MinimumCategory))
                .Append(" (").Append(MinimumCategory).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .maximumCategory      = ")
                .Append("0x").Append(HexDump.ToHex(MaximumCategory))
                .Append(" (").Append(MaximumCategory).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .majorUnitValue       = ")
                .Append("0x").Append(HexDump.ToHex(MajorUnitValue))
                .Append(" (").Append(MajorUnitValue).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .majorUnit            = ")
                .Append("0x").Append(HexDump.ToHex(MajorUnit))
                .Append(" (").Append(MajorUnit).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .minorUnitValue       = ")
                .Append("0x").Append(HexDump.ToHex(MinorUnitValue))
                .Append(" (").Append(MinorUnitValue).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .minorUnit            = ")
                .Append("0x").Append(HexDump.ToHex(MinorUnit))
                .Append(" (").Append(MinorUnit).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .baseUnit             = ")
                .Append("0x").Append(HexDump.ToHex(BaseUnit))
                .Append(" (").Append(BaseUnit).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .crossingPoint        = ")
                .Append("0x").Append(HexDump.ToHex(CrossingPoint))
                .Append(" (").Append(CrossingPoint).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .defaultMinimum           = ").Append(IsDefaultMinimum).Append('\n');
            buffer.Append("         .defaultMaximum           = ").Append(IsDefaultMaximum).Append('\n');
            buffer.Append("         .defaultMajor             = ").Append(IsDefaultMajor).Append('\n');
            buffer.Append("         .defaultMinorUnit         = ").Append(IsDefaultMinorUnit).Append('\n');
            buffer.Append("         .IsDate                   = ").Append(IsIsDate).Append('\n');
            buffer.Append("         .defaultBase              = ").Append(IsDefaultBase).Append('\n');
            buffer.Append("         .defaultCross             = ").Append(IsDefaultCross).Append('\n');
            buffer.Append("         .defaultDateSettings      = ").Append(IsDefaultDateSettings).Append('\n');

            buffer.Append("[/AXCEXT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_minimumCategory);
            out1.WriteShort(field_2_maximumCategory);
            out1.WriteShort(field_3_majorUnitValue);
            out1.WriteShort(field_4_majorUnit);
            out1.WriteShort(field_5_minorUnitValue);
            out1.WriteShort(field_6_minorUnit);
            out1.WriteShort(field_7_baseUnit);
            out1.WriteShort(field_8_crossingPoint);
            out1.WriteShort(field_9_options);
        }

        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2;
            }
            
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }


        public override object Clone()
        {
            AxisOptionsRecord rec = new AxisOptionsRecord();

            rec.field_1_minimumCategory = field_1_minimumCategory;
            rec.field_2_maximumCategory = field_2_maximumCategory;
            rec.field_3_majorUnitValue = field_3_majorUnitValue;
            rec.field_4_majorUnit = field_4_majorUnit;
            rec.field_5_minorUnitValue = field_5_minorUnitValue;
            rec.field_6_minorUnit = field_6_minorUnit;
            rec.field_7_baseUnit = field_7_baseUnit;
            rec.field_8_crossingPoint = field_8_crossingPoint;
            rec.field_9_options = field_9_options;
            return rec;
        }




        /**
         * Get the minimum category field for the AxisOptions record.
         */
        public short MinimumCategory
        {
            get
            {
                return field_1_minimumCategory;
            }
            set
            {
                this.field_1_minimumCategory = value;
            }
        }


        /**
         * Get the maximum category field for the AxisOptions record.
         */
        public short MaximumCategory
        {
            get
            {
                return field_2_maximumCategory;
            }
            set
            {
                this.field_2_maximumCategory = value;
            }
        }


        /**
         * Get the major unit value field for the AxisOptions record.
         */
        public short MajorUnitValue
        {
            get
            {
                return field_3_majorUnitValue;
            }
            set
            {
                this.field_3_majorUnitValue = value;
            }
        }

        /**
         * Get the major unit field for the AxisOptions record.
         */
        public short MajorUnit
        {
            get
            {
                return field_4_majorUnit;
            }
            set
            {
                this.field_4_majorUnit = value;
            }
        }

        /**
         * Get the minor unit value field for the AxisOptions record.
         */
        public short MinorUnitValue
        {
            get
            {
                return field_5_minorUnitValue;
            }
            set
            {
                this.field_5_minorUnitValue = value;
            }
        }

        /**
         * Get the minor unit field for the AxisOptions record.
         */
        public short MinorUnit
        {
            get
            {
                return field_6_minorUnit;
            }
            set
            {
                this.field_6_minorUnit = value;
            }
        }

        /**
         * Get the base unit field for the AxisOptions record.
         */
        public short BaseUnit
        {
            get
            {
                return field_7_baseUnit;
            }
            set
            {
                this.field_7_baseUnit = value;
            }
        }

        /**
         * Get the crossing point field for the AxisOptions record.
         */
        public short CrossingPoint
        {
            get
            {
                return field_8_crossingPoint;
            }
            set
            {
                this.field_8_crossingPoint = value;
            }
        }

        /**
         * Get the options field for the AxisOptions record.
         */
        public short Options
        {
            get
            {
                return field_9_options;
            }
            set
            {
                this.field_9_options = value;
            }
        }


        /**
         * use the default minimum category
         * @return  the default minimum field value.
         */
        public bool IsDefaultMinimum
        {
            get
            {
                return defaultMinimum.IsSet(field_9_options);
            }
            set
            {
                field_9_options = defaultMinimum.SetShortBoolean(field_9_options, value);
            }
        }

        /**
         * use the default maximum category
         * @return  the default maximum field value.
         */
        public bool IsDefaultMaximum
        {
            get
            {
                return defaultMaximum.IsSet(field_9_options);
            }
            set
            {
                field_9_options = defaultMaximum.SetShortBoolean(field_9_options, value);
            }
        }

        /**
         * use the default major unit
         * @return  the default major field value.
         */
        public bool IsDefaultMajor
        {
            get
            {
                return defaultMajor.IsSet(field_9_options);
            }
            set
            {
                field_9_options = defaultMajor.SetShortBoolean(field_9_options, value);
            }
        }

        /**
         * use the default minor unit
         * @return  the default minor unit field value.
         */
        public bool IsDefaultMinorUnit
        {
            get
            {
                return defaultMinorUnit.IsSet(field_9_options);
            }
            set
            {
                field_9_options = defaultMinorUnit.SetShortBoolean(field_9_options, value);
            }
        }

        /**
         * Sets the isDate field value.
         * this is a date axis
         */
        public void SetIsDate(bool value)
        {
            
        }

        /**
         * this is a date axis
         * @return  the isDate field value.
         */
        public bool IsIsDate
        {
            get
            {
                return isDate.IsSet(field_9_options);
            }
            set
            {
                field_9_options = isDate.SetShortBoolean(field_9_options, value);
            }
        }

        /**
         * use the default base unit
         * @return  the default base field value.
         */
        public bool IsDefaultBase
        {
            get
            {
                return defaultBase.IsSet(field_9_options);
            }
            set
            {
                field_9_options = defaultBase.SetShortBoolean(field_9_options, value);
            }
        }

        /**
         * use the default crossing point
         * @return  the default cross field value.
         */
        public bool IsDefaultCross
        {
            get
            {
                return defaultCross.IsSet(field_9_options);
            }
            set
            {
                field_9_options = defaultCross.SetShortBoolean(field_9_options, value);
            }
        }

        /**
         * use default date Setttings for this axis
         * @return  the default date Settings field value.
         */
        public bool IsDefaultDateSettings
        {
            get
            {
                return defaultDateSettings.IsSet(field_9_options);
            }
            set
            {
                field_9_options = defaultDateSettings.SetShortBoolean(field_9_options, value);
            }
        }
    }

}
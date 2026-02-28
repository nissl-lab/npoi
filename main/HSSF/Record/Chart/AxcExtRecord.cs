
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

    public enum DateUnit
    {
        Days = 0,
        Months = 1,
        Years = 2
    }
    /*
     * The axis options record provides Unit information and other various tidbits about the axis.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Andrew C. Oliver(acoliver at apache.org)
     */
    //
    /// <summary>
    /// The AxcExt record specifies additional extension properties of a date axis (section 2.2.3.6), 
    /// along with a CatSerRange record (section 2.4.39).
    /// </summary>
    public class AxcExtRecord
       : StandardRecord
    {
        public const short sid = 0x1062;
        private short field_1_catMin;
        private short field_2_catMax;
        private short field_3_catMajor;
        private short field_4_duMajor;
        private short field_5_catMinor;
        private short field_6_duMinor;
        private short field_7_duBase;
        private short field_8_catCrossDate;
        private short field_9_options;
        private BitField fAutoMin = BitFieldFactory.GetInstance(0x1);
        private BitField fAutoMax = BitFieldFactory.GetInstance(0x2);
        private BitField fAutoMajor = BitFieldFactory.GetInstance(0x4);
        private BitField fAutoMinor = BitFieldFactory.GetInstance(0x8);
        private BitField fDateAxis = BitFieldFactory.GetInstance(0x10);
        private BitField fAutoBase = BitFieldFactory.GetInstance(0x20);
        private BitField fAutoCross = BitFieldFactory.GetInstance(0x40);
        private BitField fAutoDate = BitFieldFactory.GetInstance(0x80);


        public AxcExtRecord()
        {

        }

        /*
         * Constructs a AxisOptions record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public AxcExtRecord(RecordInputStream in1)
        {
            field_1_catMin = in1.ReadShort();
            field_2_catMax = in1.ReadShort();
            field_3_catMajor = in1.ReadShort();
            field_4_duMajor = in1.ReadShort();
            field_5_catMinor = in1.ReadShort();
            field_6_duMinor = in1.ReadShort();
            field_7_duBase = in1.ReadShort();
            field_8_catCrossDate = in1.ReadShort();
            field_9_options = in1.ReadShort();

        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AXCEXT]\n");
            buffer.Append("    .catMin      = ")
                .Append("0x").Append(HexDump.ToHex(MinimumDate))
                .Append(" (").Append(MinimumDate).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .catMax      = ")
                .Append("0x").Append(HexDump.ToHex(MaximumDate))
                .Append(" (").Append(MaximumDate).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .catMajor       = ")
                .Append("0x").Append(HexDump.ToHex(MajorInterval))
                .Append(" (").Append(MajorInterval).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .duMajor            = ")
                .Append("0x").Append(HexDump.ToHex((short)MajorUnit))
                .Append(" (").Append(MajorUnit).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .catMinor       = ")
                .Append("0x").Append(HexDump.ToHex(MinorInterval))
                .Append(" (").Append(MinorInterval).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .duMinor            = ")
                .Append("0x").Append(HexDump.ToHex((short)MinorUnit))
                .Append(" (").Append(MinorUnit).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .duBase             = ")
                .Append("0x").Append(HexDump.ToHex((short)BaseUnit))
                .Append(" (").Append(BaseUnit).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .catCrossDate        = ")
                .Append("0x").Append(HexDump.ToHex(CrossDate))
                .Append(" (").Append(CrossDate).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .fAutoMin           = ").Append(IsAutoMin).Append('\n');
            buffer.Append("         .fAutoMax           = ").Append(IsAutoMax).Append('\n');
            buffer.Append("         .fAutoMajor         = ").Append(IsAutoMajor).Append('\n');
            buffer.Append("         .fAutoMinor         = ").Append(IsAutoMinor).Append('\n');
            buffer.Append("         .fDateAxis          = ").Append(IsDateAxis).Append('\n');
            buffer.Append("         .fAutoBase          = ").Append(IsAutoBase).Append('\n');
            buffer.Append("         .fAutoCross         = ").Append(IsAutoCross).Append('\n');
            buffer.Append("         .fAutoDate          = ").Append(IsAutoDate).Append('\n');

            buffer.Append("[/AXCEXT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_catMin);
            out1.WriteShort(field_2_catMax);
            out1.WriteShort(field_3_catMajor);
            out1.WriteShort(field_4_duMajor);
            out1.WriteShort(field_5_catMinor);
            out1.WriteShort(field_6_duMinor);
            out1.WriteShort(field_7_duBase);
            out1.WriteShort(field_8_catCrossDate);
            out1.WriteShort(field_9_options);
        }

        /*
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            AxcExtRecord rec = new AxcExtRecord();

            rec.field_1_catMin = field_1_catMin;
            rec.field_2_catMax = field_2_catMax;
            rec.field_3_catMajor = field_3_catMajor;
            rec.field_4_duMajor = field_4_duMajor;
            rec.field_5_catMinor = field_5_catMinor;
            rec.field_6_duMinor = field_6_duMinor;
            rec.field_7_duBase = field_7_duBase;
            rec.field_8_catCrossDate = field_8_catCrossDate;
            rec.field_9_options = field_9_options;
            return rec;
        }




        /*
         * Get the minimum category field for the AxisOptions record.
         */
        public short MinimumDate
        {
            get
            {
                return field_1_catMin;
            }
            set 
            {
                this.field_1_catMin = value;
            }
        }

        /*
         * Get the maximum category field for the AxisOptions record.
         */
        public short MaximumDate
        {
            get
            {
                return field_2_catMax;
            }
            set 
            {
                this.field_2_catMax = value;
            }
        }

        /*
         * Get the major Unit value field for the AxisOptions record.
         */
        //
        /// <summary>
        /// specifies the interval at which the major tick marks are displayed on the axis (section 2.2.3.6), 
        /// in the unit defined by duMajor.
        /// </summary>
        public short MajorInterval
        {
            get
            {
                return field_3_catMajor;
            }
            set 
            {
                this.field_3_catMajor = value;
            }
        }

        /*
         * Get the major Unit field for the AxisOptions record.
         */
        //
        /// <summary>
        /// specifies the unit of time to use for catMajor when the axis (section 2.2.3.6) is a date axis (section 2.2.3.6).
        /// If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public DateUnit MajorUnit
        {
            get
            {
                return (DateUnit)field_4_duMajor;
            }
            set
            {
                this.field_4_duMajor = (short)value;
            }
        }

        /*
         * Get the minor Unit value field for the AxisOptions record.
         */
        //
        /// <summary>
        /// specifies the interval at which the minor tick marks are displayed on the axis (section 2.2.3.6), 
        /// in a unit defined by duMinor.
        /// </summary>
        public short MinorInterval
        {
            get
            {
                return field_5_catMinor;
            }
            set
            {
                this.field_5_catMinor = value;
            }
        }

        /*
         * Get the minor Unit field for the AxisOptions record.
         */
        public DateUnit MinorUnit
        {
            get
            {
                return (DateUnit)field_6_duMinor;
            }
            set 
            {
                this.field_6_duMinor = (short)value;
            }
        }


        /*
         * Get the base Unit field for the AxisOptions record.
         */
        //
        /// <summary>
        /// specifies the smallest unit of time used by the axis (section 2.2.3.6).
        /// </summary>
        public DateUnit BaseUnit
        {
            get
            {
                return (DateUnit)field_7_duBase;
            }
            set 
            {
                this.field_7_duBase = (short)value;
            }
        }

        /*
         * Get the crossing point field for the AxisOptions record.
         */
        //
        /// <summary>
        /// specifies at which date, as a date in the date system specified by the Date1904 record (section 2.4.77), 
        /// in the units defined by duBase, the value axis (section 2.2.3.6) crosses this axis (section 2.2.3.6).
        /// </summary>
        public short CrossDate
        {
            get
            {
                return field_8_catCrossDate;
            }
            set 
            {
                this.field_8_catCrossDate = value;
            }
        }

        /*
         * Get the options field for the AxisOptions record.
         */
        public short Options
        {
            get { return field_9_options; }
            set { this.field_9_options = value; }
        }

        /*
         * use the default minimum category
         * @return  the default minimum field value.
         */
        //
        /// <summary>
        /// specifies whether MinimumDate is calculated automatically.
        /// </summary>
        public bool IsAutoMin
        {
            get
            {
                return fAutoMin.IsSet(field_9_options);
            }
            set { field_9_options = fAutoMin.SetShortBoolean(field_9_options, value); }
        }
        /*
         * use the default maximum category
         * @return  the default maximum field value.
         */
        /// <summary>
        /// specifies whether MaximumDate is calculated automatically.
        /// </summary>
        public bool IsAutoMax
        {
            get
            {
                return fAutoMax.IsSet(field_9_options);
            }
            set 
            {
                field_9_options = fAutoMax.SetShortBoolean(field_9_options, value);
            }
        }

        /*
         * use the default major Unit
         * @return  the default major field value.
         */
        public bool IsAutoMajor
        {
            get
            {
                return fAutoMajor.IsSet(field_9_options);
            }
            set 
            {
                field_9_options = fAutoMajor.SetShortBoolean(field_9_options, value);
            }
        }

        /*
         * use the default minor Unit
         * @return  the default minor Unit field value.
         */
        public bool IsAutoMinor
        {
            get
            {
                return fAutoMinor.IsSet(field_9_options);
            }
            set { field_9_options = fAutoMinor.SetShortBoolean(field_9_options, value); }
        }


        /*
         * this is a date axis
         * @return  the IsDate field value.
         */
        public bool IsDateAxis
        {
            get
            {
                return fDateAxis.IsSet(field_9_options);
            }
            set 
            {
                field_9_options = fDateAxis.SetShortBoolean(field_9_options, value);
            }
        }

        /*
         * use the default base Unit
         * @return  the default base field value.
         */
        public bool IsAutoBase
        {
            get
            {
                return fAutoBase.IsSet(field_9_options);
            }
            set { field_9_options = fAutoBase.SetShortBoolean(field_9_options, value); }
        }

        /*
         * use the default crossing point
         * @return  the default cross field value.
         */
        public bool IsAutoCross
        {
            get
            {
                return fAutoCross.IsSet(field_9_options);
            }
            set 
            {
                field_9_options = fAutoCross.SetShortBoolean(field_9_options, value);
            }
        }
        /*
         * use default date Setttings for this axis
         * @return  the default date Settings field value.
         */
        public bool IsAutoDate
        {
            get
            {
                return fAutoDate.IsSet(field_9_options);
            }
            set { field_9_options = fAutoDate.SetShortBoolean(field_9_options, value); }
        }


    }
}



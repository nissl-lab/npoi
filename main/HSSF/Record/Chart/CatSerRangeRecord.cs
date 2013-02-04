
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

    /*
     * This record refers to a category or series axis and is used to specify label/tickmark frequency.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    //
    /// <summary>
    /// specifies the properties of a category (3) axis, a date axis, or a series axis.
    /// </summary>
    public class CatSerRangeRecord
       : StandardRecord
    {
        public const short sid = 0x1020;
        private short field_1_catCross;
        private short field_2_catLabel;
        private short field_3_catMark;
        private short field_4_options;
        private BitField fBetween = BitFieldFactory.GetInstance(0x1);
        private BitField fMaxCross = BitFieldFactory.GetInstance(0x2);
        private BitField fReverse = BitFieldFactory.GetInstance(0x4);


        public CatSerRangeRecord()
        {

        }

        /**
         * Constructs a CategorySeriesAxis record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public CatSerRangeRecord(RecordInputStream in1)
        {
            field_1_catCross = in1.ReadShort();
            field_2_catLabel = in1.ReadShort();
            field_3_catMark = in1.ReadShort();
            field_4_options = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CATSERRANGE]\n");
            buffer.Append("    .catCross        = ")
                .Append("0x").Append(HexDump.ToHex(CrossPoint))
                .Append(" (").Append(CrossPoint).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .catLabel       = ")
                .Append("0x").Append(HexDump.ToHex(LabelInterval))
                .Append(" (").Append(LabelInterval).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .catMark    = ")
                .Append("0x").Append(HexDump.ToHex(MarkInterval))
                .Append(" (").Append(MarkInterval).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .fBetween        = ").Append(IsBetween).Append('\n');
            buffer.Append("         .fMaxCross       = ").Append(IsMaxCross).Append('\n');
            buffer.Append("         .fReverse        = ").Append(IsReverse).Append('\n');

            buffer.Append("[/CATSERRANGE]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_catCross);
            out1.WriteShort(field_2_catLabel);
            out1.WriteShort(field_3_catMark);
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
            CatSerRangeRecord rec = new CatSerRangeRecord();

            rec.field_1_catCross = field_1_catCross;
            rec.field_2_catLabel = field_2_catLabel;
            rec.field_3_catMark = field_3_catMark;
            rec.field_4_options = field_4_options;
            return rec;
        }




        /*
         * Get the crossing point field for the CategorySeriesAxis record.
         */
        //
        /// <summary>
        /// specifies where the value axis crosses this axis, based on the following table.
        /// If fMaxCross is set to 1, the value this field MUST be ignored.
        /// Category (3) axis   This field specifies the category (3) at which the value axis crosses. 
        ///                     For example, if this field is 2, the value axis crosses this axis at the second category (3) 
        ///                     on this axis. MUST be greater than or equal to 1 and less than or equal to 31999.
        /// Series axis         MUST be 0.
        /// Date axis           catCross MUST be equal to the value given by the following formula:
        ///                     catCross = catCrossDate ¨C catMin + 1
        ///                     Where catCrossDate is the catCrossDate field of the AxcExt record 
        ///                     and catMin is the catMin field of the AxcExt record.
        /// </summary>
        public short CrossPoint
        {
            get
            {
                return field_1_catCross;
            }
            set
            {
                this.field_1_catCross = value;
            }
        }

        /*
         * Get the label frequency field for the CategorySeriesAxis record.
         */
        //
        /// <summary>
        /// specifies the interval between axis labels on this axis. MUST be greater than or equal to 1 and 
        /// less than or equal to 31999. MUST be ignored for a date axis.
        /// </summary>
        public short LabelInterval
        {
            get
            {
                return field_2_catLabel;
            }
            set
            {
                this.field_2_catLabel = value;
            }
        }

        /*
         * Get the tick mark frequency field for the CategorySeriesAxis record.
         */
        
        //
        /// <summary>
        /// specifies the interval at which major tick marks and minor tick marks are displayed on the axis. 
        /// Major tick marks and minor tick marks that would have been visible are hidden unless they are 
        /// located at a multiple of this field.
        /// </summary>
        public short MarkInterval
        {
            get
            {
                return field_3_catMark;
            }
            set
            {
                this.field_3_catMark = value;
            }
        }


        /*
         * Get the options field for the CategorySeriesAxis record.
         */
        public short Options
        {
            get { return field_4_options; }
            set { this.field_4_options = value; }
        }

        /*
         * Set true to indicate axis crosses between categories and false to cross axis midway
         * @return  the value axis crossing field value.
         */
        //
        /// <summary>
        /// specifies whether the value axis crosses this axis between major tick marks. MUST be a value from to following table:
        /// 0  The value axis crosses this axis on a major tick mark.
        /// 1  The value axis crosses this axis between major tick marks.
        /// </summary>
        public bool IsBetween
        {
            get
            {
                return fBetween.IsSet(field_4_options);
            }
            set
            {
                field_4_options = fBetween.SetShortBoolean(field_4_options, value);
            }
        }

        /*
         * axis crosses at the far right
         * @return  the crosses far right field value.
         */
        //
        /// <summary>
        /// specifies whether the value axis crosses this axis at the last category (3), the last series, 
        /// or the maximum date. MUST be a value from the following table:
        /// 0  The value axis crosses this axis at the value specified by catCross.
        /// 1  The value axis crosses this axis at the last category (3), the last series, or the maximum date.
        /// </summary>
        public bool IsMaxCross
        {
            get
            {
                return fMaxCross.IsSet(field_4_options);
            }
            set
            {
                field_4_options = fMaxCross.SetShortBoolean(field_4_options, value);
            }
        }

        /*
         * categories are Displayed in reverse order
         * @return  the reversed field value.
         */
        //
        /// <summary>
        /// specifies whether the axis is displayed in reverse order. MUST be a value from the following table:
        /// 0  The axis is displayed in order.
        /// 1  The axis is display in reverse order.
        /// </summary>
        public bool IsReverse
        {
            get
            {
                return fReverse.IsSet(field_4_options);
            }
            set
            {
                field_4_options = fReverse.SetShortBoolean(field_4_options, value);
            }
        }


    }
}



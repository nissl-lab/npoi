
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
     * The frame record indicates whether there is a border around the Displayed text of a chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class FrameRecord
       : StandardRecord
    {
        public const short sid = 0x1032;
        private short field_1_borderType;
        public const short BORDER_TYPE_REGULAR = 0;
        public const short BORDER_TYPE_SHADOW = 1;
        private short field_2_options;
        private BitField autoSize = BitFieldFactory.GetInstance(0x1);
        private BitField autoPosition = BitFieldFactory.GetInstance(0x2);


        public FrameRecord()
        {

        }

        /**
         * Constructs a Frame record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public FrameRecord(RecordInputStream in1)
        {
            field_1_borderType = in1.ReadShort();
            field_2_options = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FRAME]\n");
            buffer.Append("    .borderType           = ")
                .Append("0x").Append(HexDump.ToHex(BorderType))
                .Append(" (").Append(BorderType).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options              = ")
                .Append("0x").Append(HexDump.ToHex(Options))
                .Append(" (").Append(Options).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .autoSize                 = ").Append(IsAutoSize).Append('\n');
            buffer.Append("         .autoPosition             = ").Append(IsAutoPosition).Append('\n');

            buffer.Append("[/FRAME]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_borderType);
            out1.WriteShort(field_2_options);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            FrameRecord rec = new FrameRecord();

            rec.field_1_borderType = field_1_borderType;
            rec.field_2_options = field_2_options;
            return rec;
        }




        /**
         * Get the border type field for the Frame record.
         *
         * @return  One of 
         *        BORDER_TYPE_REGULAR
         *        BORDER_TYPE_SHADOW
         */
        public short BorderType
        {
            get
            {
                return field_1_borderType;
            }
            set 
            {
                this.field_1_borderType = value;
            }
        }
        /**
         * Get the options field for the Frame record.
         */
        public short Options
        {
            get { return field_2_options; }
            set { this.field_2_options = value; }
        }

        /**
         * excel calculates the size automatically if true
         * @return  the auto size field value.
         */
        public bool IsAutoSize
        {
            get
            {
                return autoSize.IsSet(field_2_options);
            }
            set
            {
                field_2_options = autoSize.SetShortBoolean(field_2_options, value);
            }
        }


        /**
         * excel calculates the position automatically
         * @return  the auto position field value.
         */
        public bool IsAutoPosition
        {
            get
            {
                return autoPosition.IsSet(field_2_options);
            }
            set 
            {
                field_2_options = autoPosition.SetShortBoolean(field_2_options, value);
            }
        }


    }
}


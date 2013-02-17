
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
     * The axis record defines the type of an axis.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class AxisRecord
       : StandardRecord
    {
        public const short sid = 0x101d;
        private short field_1_axisType;
        public const short AXIS_TYPE_CATEGORY_OR_X_AXIS = 0;
        public const short AXIS_TYPE_VALUE_AXIS = 1;
        public const short AXIS_TYPE_SERIES_AXIS = 2;
        private int field_2_reserved1;
        private int field_3_reserved2;
        private int field_4_reserved3;
        private int field_5_reserved4;


        public AxisRecord()
        {

        }

        /**
         * Constructs a Axis record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public AxisRecord(RecordInputStream in1)
        {
            field_1_axisType = in1.ReadShort();
            field_2_reserved1 = in1.ReadInt();
            field_3_reserved2 = in1.ReadInt();
            field_4_reserved3 = in1.ReadInt();
            field_5_reserved4 = in1.ReadInt();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AXIS]\n");
            buffer.Append("    .axisType             = ")
                .Append("0x").Append(HexDump.ToHex(AxisType))
                .Append(" (").Append(AxisType).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .reserved1            = ")
                .Append("0x").Append(HexDump.ToHex(Reserved1))
                .Append(" (").Append(Reserved1).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .reserved2            = ")
                .Append("0x").Append(HexDump.ToHex(Reserved2))
                .Append(" (").Append(Reserved2).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .reserved3            = ")
                .Append("0x").Append(HexDump.ToHex(Reserved3))
                .Append(" (").Append(Reserved3).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .reserved4            = ")
                .Append("0x").Append(HexDump.ToHex(Reserved4))
                .Append(" (").Append(Reserved4).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/AXIS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {            
            out1.WriteShort(field_1_axisType);
            out1.WriteInt(field_2_reserved1);
            out1.WriteInt(field_3_reserved2);
            out1.WriteInt(field_4_reserved3);
            out1.WriteInt(field_5_reserved4);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 4 + 4 + 4 + 4; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            AxisRecord rec = new AxisRecord();

            rec.field_1_axisType = field_1_axisType;
            rec.field_2_reserved1 = field_2_reserved1;
            rec.field_3_reserved2 = field_3_reserved2;
            rec.field_4_reserved3 = field_4_reserved3;
            rec.field_5_reserved4 = field_5_reserved4;
            return rec;
        }




        /**
         * Get the axis type field for the Axis record.
         *
         * @return  One of 
         *        AXIS_TYPE_CATEGORY_OR_X_AXIS
         *        AXIS_TYPE_VALUE_AXIS
         *        AXIS_TYPE_SERIES_AXIS
         */
        public short AxisType
        {
            get { return field_1_axisType; }
            set { this.field_1_axisType = value; }
        }

        /**
         * Get the reserved1 field for the Axis record.
         */
        public int Reserved1
        {
            get { return field_2_reserved1; }
            set { this.field_2_reserved1 = value; }
        }

        /**
         * Get the reserved2 field for the Axis record.
         */
        public int Reserved2
        {
            get { return field_3_reserved2; }
            set { this.field_3_reserved2 = value; }
        }

        /**
         * Get the reserved3 field for the Axis record.
         */
        public int Reserved3
        {
            get { return field_4_reserved3; }
            set { this.field_4_reserved3 = value; }
        }

        /**
         * Get the reserved4 field for the Axis record.
         */
        public int Reserved4
        {
            get { return field_5_reserved4; }
            set { this.field_5_reserved4 = value; }
        }


    }
}




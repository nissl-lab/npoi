
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
     * The axis size and location
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class AxisParentRecord
       : StandardRecord
    {
        public const short sid = 0x1041;
        private short field_1_axisType;
        public const short AXIS_TYPE_MAIN = 0;
        public const short AXIS_TYPE_SECONDARY = 1;
        private int field_2_x;
        private int field_3_y;
        private int field_4_width;
        private int field_5_height;


        public AxisParentRecord()
        {

        }

        /**
         * Constructs a AxisParent record and s its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public AxisParentRecord(RecordInputStream in1)
        {

            field_1_axisType = in1.ReadShort();
            field_2_x = in1.ReadInt();
            field_3_y = in1.ReadInt();
            field_4_width = in1.ReadInt();
            field_5_height = in1.ReadInt();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[AXISPARENT]\n");
            buffer.Append("    .axisType             = ")
                .Append("0x").Append(HexDump.ToHex(AxisType))
                .Append(" (").Append(AxisType).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .x                    = ")
                .Append("0x").Append(HexDump.ToHex(X))
                .Append(" (").Append(X).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .y                    = ")
                .Append("0x").Append(HexDump.ToHex(Y))
                .Append(" (").Append(Y).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .width                = ")
                .Append("0x").Append(HexDump.ToHex(Width))
                .Append(" (").Append(Width).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .height               = ")
                .Append("0x").Append(HexDump.ToHex(Height))
                .Append(" (").Append(Height).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/AXISPARENT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_axisType);
            out1.WriteInt(field_2_x);
            out1.WriteInt(field_3_y);
            out1.WriteInt(field_4_width);
            out1.WriteInt(field_5_height);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return  2 + 4 + 4 + 4 + 4; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            AxisParentRecord rec = new AxisParentRecord();

            rec.field_1_axisType = field_1_axisType;
            rec.field_2_x = field_2_x;
            rec.field_3_y = field_3_y;
            rec.field_4_width = field_4_width;
            rec.field_5_height = field_5_height;
            return rec;
        }




        /**
         *  the axis type field for the AxisParent record.
         *
         * @return  One of 
         *        AXIS_TYPE_MAIN
         *        AXIS_TYPE_SECONDARY
         */
        public short AxisType
        {
            get { return field_1_axisType; }
            set { this.field_1_axisType = value; }
        }

        /**
         *  the x field for the AxisParent record.
         */
        public int X
        {
            get { return field_2_x; }
            set { this.field_2_x = value; }
        }

        /**
         *  the y field for the AxisParent record.
         */
        public int Y
        {
            get { return field_3_y; }
            set { this.field_3_y = value; }
        }

        /**
         *  the width field for the AxisParent record.
         */
        public int Width
        {
            get { return field_4_width; }
            set { this.field_4_width = value; }
        }

        /**
         *  the height field for the AxisParent record.
         */
        public int Height
        {
            get { return field_5_height; }
            set { this.field_5_height = value; }
        }

    }
}



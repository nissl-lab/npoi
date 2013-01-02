
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
     * The chart record is used to define the location and size of a chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class ChartRecord
       : StandardRecord
    {
        public const short sid = 0x1002;
        private int field_1_x;
        private int field_2_y;
        private int field_3_width;
        private int field_4_height;


        public ChartRecord()
        {

        }

        /**
         * Constructs a Chart record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public ChartRecord(RecordInputStream in1)
        {
            field_1_x = in1.ReadInt();
            field_2_y = in1.ReadInt();
            field_3_width = in1.ReadInt();
            field_4_height = in1.ReadInt();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CHART]\n");
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

            buffer.Append("[/CHART]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_x);
            out1.WriteInt(field_2_y);
            out1.WriteInt(field_3_width);
            out1.WriteInt(field_4_height);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return  4 + 4 + 4 + 4; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            ChartRecord rec = new ChartRecord();

            rec.field_1_x = field_1_x;
            rec.field_2_y = field_2_y;
            rec.field_3_width = field_3_width;
            rec.field_4_height = field_4_height;
            return rec;
        }




        /**
         * Get the x field for the Chart record.
         */
        public int X
        {
            get
            {
                return field_1_x;
            }
            set 
            {
                this.field_1_x = value;
            }
        }

        /**
         * Get the y field for the Chart record.
         */
        public int Y
        {
            get
            {
                return field_2_y;
            }
            set 
            {
                this.field_2_y = value;
            }
        }

        /**
         * Get the width field for the Chart record.
         */
        public int Width
        {
            get { return field_3_width; }
            set { this.field_3_width = value; }
        }


        /**
         * Get the height field for the Chart record.
         */
        public int Height
        {
            get{return field_4_height;}
            set{this.field_4_height = value;}
        }

    }
}



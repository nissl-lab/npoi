
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


namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * Class ChartFormatRecord
     *
     *
     * @author Glen Stampoultzis (glens at apache.org)
     * @version %I%, %G%
     */

    public class ChartFormatRecord
       : StandardRecord
    {
        public const short sid = 0x1014;

        // ignored?
        private int field1_x_position;   // lower left
        private int field2_y_position;   // lower left
        private int field3_width;
        private int field4_height;
        
        private short field5_grbit;
        private BitField varyDisplayPattern = BitFieldFactory.GetInstance(0x01);

        private short field6_icrt;

        public ChartFormatRecord()
        {
        }

        /**
         * Constructs a ChartFormatRecord record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public ChartFormatRecord(RecordInputStream in1)
        {
            field1_x_position = in1.ReadInt();
            field2_y_position = in1.ReadInt();
            field3_width = in1.ReadInt();
            field4_height = in1.ReadInt();
            field5_grbit = in1.ReadShort();
            field6_icrt = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CHARTFORMAT]\n");
            buffer.Append("    .xPosition       = ").Append(XPosition)
                .Append("\n");
            buffer.Append("    .yPosition       = ").Append(YPosition)
                .Append("\n");
            buffer.Append("    .width           = ").Append(Width)
                .Append("\n");
            buffer.Append("    .height          = ").Append(Height)
                .Append("\n");
            buffer.Append("    .grBit           = ")
                .Append(StringUtil.ToHexString(field5_grbit)).Append("\n");
            buffer.Append("[/CHARTFORMAT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(XPosition);
            out1.WriteInt(YPosition);
            out1.WriteInt(Width);
            out1.WriteInt(Height);
            out1.WriteShort(field5_grbit);
            out1.WriteShort(field6_icrt);
        }


        public override object Clone()
        {
            ChartFormatRecord r = new ChartFormatRecord();
            r.Height = this.Height;
            r.Icrt = this.Icrt;
            r.VaryDisplayPattern = this.VaryDisplayPattern;
            r.Width = this.Width;
            r.XPosition = this.XPosition;
            r.YPosition = this.YPosition;
            return r;
        }

        protected override int DataSize
        {
            get { return 20; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public int XPosition
        {
            get
            {
                return field1_x_position;
            }
            set
            {
                this.field1_x_position = value;
            }
        }

        public int YPosition
        {
            get
            {
                return field2_y_position;
            }
            set 
            {
                this.field2_y_position = value;
            }
        }

        public int Width
        {
            get { return field3_width; }
            set { this.field3_width = value; }
        }

        public int Height
        {
            get { return field4_height; }
            set { this.field4_height = value; }
        }

        public short Icrt
        {
            get { return field6_icrt; }
            set { this.field6_icrt=value; }
        }

        public bool VaryDisplayPattern
        {
            get
            {
                return varyDisplayPattern.IsSet(field5_grbit);
            }
            set 
            {
                field5_grbit = varyDisplayPattern.SetShortBoolean(field5_grbit,
                        value);           
            }
        }
    }
}
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
using System;
using System.Text;

using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    public class Chart3DBarShapeRecord:StandardRecord
    {
        public const short sid = 4191;

        byte field_1_riser = 0;
        byte field_2_taper = 0;

        public Chart3DBarShapeRecord()
        { 
        
        }

        public Chart3DBarShapeRecord(RecordInputStream in1)
        {
            field_1_riser = (byte)in1.ReadByte();
            field_2_taper = (byte)in1.ReadByte();
        }

        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[Chart3DBarShape]\n");
            buffer.Append("    .axisType             = ")
                .Append("0x").Append(HexDump.ToHex(Riser))
                .Append(" (").Append(Riser).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .x                    = ")
                .Append("0x").Append(HexDump.ToHex(Taper))
                .Append(" (").Append(Taper).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/Chart3DBarShape]\n");
            return buffer.ToString();
        }
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(field_1_riser);
            out1.WriteByte(field_2_taper);
        }

        protected override int DataSize
        {
            get { return 1 + 1; }
        }

        public override object Clone()
        {
            Chart3DBarShapeRecord record = new Chart3DBarShapeRecord();
            record.Riser = this.Riser;
            record.Taper = this.Taper;
            return record;
        }

        public override short Sid
        {
            get { return sid; }
        }
        /// <summary>
        /// the shape of the base of the data points in a bar or column chart group. 
        /// MUST be a value from the following table
        /// 0x00      The base of the data point is a rectangle.
        /// 0x01      The base of the data point is an ellipse.
        /// </summary>
        public byte Riser
        {
            get { return field_1_riser; }
            set { field_1_riser = value; }
        }

        /// <summary>
        /// how the data points in a bar or column chart group taper from base to tip. 
        /// MUST be a value from the following
        /// 0x00    The data points of the bar or column chart group do not taper. 
        ///         The shape at the maximum value of the data point is the same as the shape at the base.:
        /// 0x01    The data points of the bar or column chart group taper to a point at the maximum value of each data point.
        /// 0x02    The data points of the bar or column chart group taper towards a projected point at the position of 
        ///         the maximum value of all of the data points in the chart group, but are clipped at the value of each data point.
        /// </summary>
        public byte Taper
        {
            get { return field_2_taper; }
            set { field_2_taper = value; }
        }
    }
}

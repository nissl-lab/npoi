
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

using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// pecifies positioning mode for position information saved in a Pos record.
    /// </summary>
    public enum PositionMode : short
    { 
        /// <summary>
        /// Relative position to the chart, in points.
        /// </summary>
        MDFX = 0x0000,
        /// <summary>
        /// Absolute width and height in points. It can only be applied to the mdBotRt field of Pos.
        /// </summary>
        MDABS = 0x0001,
        /// <summary>
        /// Owner of Pos determines how to interpret the position data.
        /// </summary>
        MDPARENT = 0x0002,
        /// <summary>
        /// Offset to default position, in 1/1000th of the plot area size.
        /// </summary>
        MDKTH = 0x0003,
        /// <summary>
        /// Relative position to the chart, in SPRC.
        /// </summary>
        MDCHART = 0x0005
    }
    /// <summary>
    /// specifies the size and position for a legend, an attached label, or the plot area, as specified by the primary axis group.
    /// </summary>
    public class PosRecord : StandardRecord
    {

        public PosRecord()
        { 
        
        }

        public const short sid = 0x104F;

        //the positioning mode for the upper-left corner
        private short mdTopLt;
        //the positioning mode for the lower-right corner
        private short mdBotRt;
        private short x1;
        private short y1;
        private short x2;
        private short y2;

        public PosRecord(RecordInputStream in1)
        {
            mdTopLt = in1.ReadShort();
            mdBotRt = in1.ReadShort();
            x1 = in1.ReadShort();
            in1.ReadShort();    //unused1
            y1 = in1.ReadShort();
            in1.ReadShort();    //unused2
            x2 = in1.ReadShort();
            in1.ReadShort();    //unused3
            y2 = in1.ReadShort();
            in1.ReadShort();    //unused4
        }

        protected override int DataSize
        {
            get { return 20; }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(mdTopLt);
            out1.WriteShort(mdBotRt);
            out1.WriteShort(x1);
            out1.WriteShort(0);
            out1.WriteShort(y1);
            out1.WriteShort(0);
            out1.WriteShort(x2);
            out1.WriteShort(0);
            out1.WriteShort(y2);
            out1.WriteShort(0);
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override object Clone()
        {
            PosRecord r = new PosRecord();
            r.MdBotRt = this.MdBotRt;
            r.MDTopLt = this.MDTopLt;
            r.X1 = this.X1;
            r.X2 = this.X2;
            r.Y1 = this.Y1;
            r.Y2 = this.Y2;
            return r;
        }
        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[POS]\n");
            buffer.Append("mdTopLt       = ").Append(HexDump.ShortToHex(mdTopLt)).Append("\n");
            buffer.Append("mdBotRt       = ").Append(HexDump.ShortToHex(mdTopLt)).Append("\n");
            buffer.Append("x1            = ").Append(HexDump.ShortToHex(x1)).Append("\n");
            buffer.Append("x2            = ").Append(HexDump.ShortToHex(x2)).Append("\n");
            buffer.Append("y1            = ").Append(HexDump.ShortToHex(y1)).Append("\n");
            buffer.Append("y2            = ").Append(HexDump.ShortToHex(y2)).Append("\n");
            buffer.Append("[/POS]\n");
            return buffer.ToString();
        }
        /// <summary>
        /// specifies the positioning mode for the upper-left corner of a legend, an attached label, or the plot area.
        /// </summary>
        public PositionMode MDTopLt
        {
            get 
            {
                return (PositionMode)mdTopLt;
            }
            set 
            {
                mdTopLt = (short)value;
            }
        }
        /// <summary>
        /// specifies the positioning mode for the lower-right corner of a legend, an attached label, or the plot area
        /// </summary>
        public PositionMode MdBotRt
        {
            get
            {
                return (PositionMode)mdBotRt;
            }
            set {
                mdBotRt = (short)value;
            }
        }
        /// <summary>
        /// specifies a position. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
        /// </summary>
        public short X1
        {
            get { return x1; }
            set { x1 = value; }
        }
        /// <summary>
        /// specifies a width. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
        /// </summary>
        public short X2
        {
            get { return x2; }
            set { x2 = value; }            
        }
        /// <summary>
        /// specifies a position. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
        /// </summary>
        public short Y1
        {
            get { return y1; }
            set { y1 = value; }
        }
        /// <summary>
        /// specifies a height. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
        /// </summary>
        public short Y2
        {
            get { return y2; }
            set { y2 = value; }
        }
    }
}

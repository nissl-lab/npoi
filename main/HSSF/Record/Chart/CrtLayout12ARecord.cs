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
    /// <summary>
    /// The CrtLayout12A record specifies layout information for a plot area.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class CrtLayout12ARecord : StandardRecord
    {
        public const short sid = 0x8A7;

        private FrtHeader frtHeader;
        private int field_1_dwCheckSum;
        private short field_2_option;
        private short field_3_xTL;
        private short field_4_yTL;
        private short field_5_xBR;
        private short field_6_yBR;
        private short field_7_wXMode;
        private short field_8_wYMode;
        private short field_9_wWidthMode;
        private short field_10_wHeightMode;
        private double field_11_x;
        private double field_12_y;
        private double field_13_dx;
        private double field_14_dy;
        private short reserved2;

        private BitField fLayoutTargetInner = BitFieldFactory.GetInstance(0x1);

        public CrtLayout12ARecord()
        {
            frtHeader.rt = (ushort)sid;
            frtHeader.grbitFrt = 0;
        }

        public override object Clone()
        {
            CrtLayout12ARecord record = new CrtLayout12ARecord();
            record.IsLayoutTargetInner = this.IsLayoutTargetInner;
            record.CheckSum = this.CheckSum;
            record.DX = this.DX;
            record.DY = this.DY;
            record.HeightMode = this.HeightMode;
            record.WidthMode = this.WidthMode;
            record.X = this.X;
            record.XMode = this.XMode;
            record.Y = this.Y;
            record.YMode = this.YMode;
            record.XTL = this.XTL;
            record.YTL = this.YTL;
            record.XBR = this.XBR;
            record.YBR = this.YBR;
            return record;
        }
        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[CRTLAYOUT12A]").AppendLine()
                .Append("   .rt               =").Append(HexDump.ToHex(frtHeader.rt)).Append("(").Append(frtHeader.rt).AppendLine(")")
                .Append("   .grbit            =").Append(HexDump.ToHex(frtHeader.grbitFrt)).Append("(").Append(frtHeader.grbitFrt).AppendLine(")")
                .Append("   .reserved         =").Append(HexDump.ToHex(0)).Append("(").Append(0).AppendLine(")")
                .Append("   .dwCheckSum       =").Append(HexDump.ToHex(field_1_dwCheckSum)).Append("(").Append(field_1_dwCheckSum).AppendLine(")")
                .Append("   .option           =").Append(HexDump.ToHex(field_2_option)).Append("(").Append(field_2_option).AppendLine(")")
                .Append("       .fLayoutTargetInner =").Append(this.IsLayoutTargetInner).AppendLine()
                .Append("   .xTL              =").Append(HexDump.ToHex(field_3_xTL)).Append("(").Append(field_3_xTL).AppendLine(")")
                .Append("   .yTL              =").Append(HexDump.ToHex(field_4_yTL)).Append("(").Append(field_4_yTL).AppendLine(")")
                .Append("   .xBR              =").Append(HexDump.ToHex(field_5_xBR)).Append("(").Append(field_5_xBR).AppendLine(")")
                .Append("   .yBR              =").Append(HexDump.ToHex(field_6_yBR)).Append("(").Append(field_6_yBR).AppendLine(")")
                .Append("   .wXMode           =").Append(HexDump.ToHex(field_7_wXMode)).Append("(").Append(field_7_wXMode).AppendLine(")")
                .Append("   .wYMode           =").Append(HexDump.ToHex(field_8_wYMode)).Append("(").Append(field_8_wYMode).AppendLine(")")
                .Append("   .wWidthMode       =").Append(HexDump.ToHex(field_9_wWidthMode)).Append("(").Append(field_9_wWidthMode).AppendLine(")")
                .Append("   .wHeightMode      =").Append(HexDump.ToHex(field_10_wHeightMode)).Append("(").Append(field_10_wHeightMode).AppendLine(")")
                .Append("   .x                =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_11_x))).Append("(").Append(field_11_x).AppendLine(")")
                .Append("   .y                =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_12_y))).Append("(").Append(field_12_y).AppendLine(")")
                .Append("   .dx               =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_13_dx))).Append("(").Append(field_13_dx).AppendLine(")")
                .Append("   .dy               =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_14_dy))).Append("(").Append(field_14_dy).AppendLine(")")
                .AppendLine("[/CRTLAYOUT12A]");
            return buffer.ToString();
        }
        public CrtLayout12ARecord(RecordInputStream ris)
        {
            frtHeader.rt = (ushort)ris.ReadUShort();
            frtHeader.grbitFrt = (ushort)ris.ReadUShort();
            ris.ReadLong();
            field_1_dwCheckSum = ris.ReadInt();
            field_2_option = ris.ReadShort();
            field_3_xTL = ris.ReadShort();
            field_4_yTL = ris.ReadShort();
            field_5_xBR = ris.ReadShort();
            field_6_yBR = ris.ReadShort();
            field_7_wXMode = ris.ReadShort();
            field_8_wYMode = ris.ReadShort();
            field_9_wWidthMode = ris.ReadShort();
            field_10_wHeightMode = ris.ReadShort();
            field_11_x = ris.ReadDouble();
            field_12_y = ris.ReadDouble();
            field_13_dx = ris.ReadDouble();
            field_14_dy = ris.ReadDouble();
            reserved2 = ris.ReadShort();
        }

        protected override int DataSize
        {
            get { return 12 + 4 + 9 * 2 + 8 * 4 + 2; }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteShort(frtHeader.rt);
            out1.WriteShort(frtHeader.grbitFrt);
            out1.WriteLong(frtHeader.reserved);
            out1.WriteInt(field_1_dwCheckSum);
            out1.WriteShort(field_2_option);
            out1.WriteShort(field_3_xTL);
            out1.WriteShort(field_4_yTL);
            out1.WriteShort(field_5_xBR);
            out1.WriteShort(field_6_yBR);
            out1.WriteShort(field_7_wXMode);
            out1.WriteShort(field_8_wYMode);
            out1.WriteShort(field_9_wWidthMode);
            out1.WriteShort(field_10_wHeightMode);
            out1.WriteDouble(field_11_x);
            out1.WriteDouble(field_12_y);
            out1.WriteDouble(field_13_dx);
            out1.WriteDouble(field_14_dy);
            out1.WriteShort(reserved2);
        }

        public override short Sid
        {
            get { return sid; }
        }

        /// <summary>
        /// specifies the type of plot area for the layout target.
        /// false  Outer plot area - The bounding rectangle that includes the axis labels, axis titles, data table (2) and plot area of the chart.
        /// true   Inner plot area – The rectangle bounded by the chart axes.
        /// </summary>
        public bool IsLayoutTargetInner
        {
            get { return fLayoutTargetInner.IsSet(field_2_option); }
            set { field_2_option = fLayoutTargetInner.SetShortBoolean(field_2_option, value); }
        }

        /// <summary>
        /// specifies the checksum
        /// </summary>
        public int CheckSum
        {
            get { return field_1_dwCheckSum; }
            set { field_1_dwCheckSum = value; }
        }
        /// <summary>
        /// specifies the horizontal offset of the plot area’s upper-left corner, relative to the upper-left corner of the chart area
        /// </summary>
        public short XTL
        {
            get { return field_3_xTL; }
            set { field_3_xTL = value; }
        }
        /// <summary>
        /// specifies the vertical offset of the plot area’s upper-left corner, relative to the upper-left corner of the chart area
        /// </summary>
        public short YTL
        {
            get { return field_4_yTL; }
            set { field_4_yTL = value; }
        }
        /// <summary>
        /// specifies the width of the plot area
        /// </summary>
        public short XBR
        {
            get { return field_5_xBR; }
            set { field_5_xBR = value; }
        }
        /// <summary>
        /// specifies the height of the plot area
        /// </summary>
        public short YBR
        {
            get { return field_6_yBR; }
            set { field_6_yBR = value; }
        }
        /// <summary>
        /// A CrtLayout12Mode structure that specifies the meaning of x.
        /// </summary>
        public CrtLayout12Mode XMode
        {
            get { return (CrtLayout12Mode)field_7_wXMode; }
            set { field_7_wXMode = (short)value; }
        }
        /// <summary>
        /// A CrtLayout12Mode structure that specifies the meaning of y.
        /// </summary>
        public CrtLayout12Mode YMode
        {
            get { return (CrtLayout12Mode)field_8_wYMode; }
            set { field_8_wYMode = (short)value; }
        }
        /// <summary>
        /// A CrtLayout12Mode structure that specifies the meaning of dx.
        /// </summary>
        public CrtLayout12Mode WidthMode
        {
            get { return (CrtLayout12Mode)field_9_wWidthMode; }
            set { field_9_wWidthMode = (short)value; }
        }
        /// <summary>
        /// A CrtLayout12Mode structure that specifies the meaning of dy.
        /// </summary>
        public CrtLayout12Mode HeightMode
        {
            get { return (CrtLayout12Mode)field_10_wHeightMode; }
            set { field_10_wHeightMode = (short)value; }
        }
        /// <summary>
        /// An Xnum (section 2.5.342) value that specifies a horizontal offset. The meaning is determined by wXMode.
        /// </summary>
        public double X
        {
            get { return field_11_x; }
            set { field_11_x = value; }
        }
        /// <summary>
        /// An Xnum value that specifies a vertical offset. The meaning is determined by wYMode.
        /// </summary>
        public double Y
        {
            get { return field_12_y; }
            set { field_12_y = value; }
        }
        /// <summary>
        /// An Xnum value that specifies a width or an horizontal offset. The meaning is determined by wWidthMode.
        /// </summary>
        public double DX
        {
            get { return field_13_dx; }
            set { field_13_dx = value; }
        }
        /// <summary>
        /// An Xnum value that specifies a height or an vertical offset. The meaning is determined by wHeightMode.
        /// </summary>
        public double DY
        {
            get { return field_14_dy; }
            set { field_14_dy = value; }
        }
    }
}

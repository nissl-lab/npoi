using System;
using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{

    /// <summary>
    /// The CrtLayout12Mode specifies a layout mode. Each layout mode specifies a different 
    /// meaning of the x, y, dx, and dy fields of CrtLayout12 and CrtLayout12A.
    /// </summary>
    public enum CrtLayout12Mode : short
    {
        /// <summary>
        /// Position and dimension (2) are determined by the application. x, y, dx and dy MUST be ignored.
        /// </summary>
        L12MAUTO = 0,
        /// <summary>
        /// x and y specify the offset of the top left corner, relative to its default position, 
        /// as a fraction of the chart area. MUST be greater than or equal to -1.0 and MUST be 
        /// less than or equal to 1.0. dx and dy specify the width and height, as a fraction of 
        /// the chart area, MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0.
        /// </summary>
        L12MFACTOR = 1,
        /// <summary>
        /// x and y specify the offset of the upper-left corner; dx and dy specify the offset of the bottom-right corner. 
        /// x, y, dx and dy are specified relative to the upper-left corner of the chart area as a fraction of the chart area. 
        /// x, y, dx and dy MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0.
        /// </summary>
        L12MEDGE = 2
    }

    /// <summary>
    /// The CrtLayout12 record specifies the layout information for attached label, when contained 
    /// in the sequence of records that conforms to the ATTACHEDLABEL rule, 
    /// or legend, when contained in the sequence of records that conforms to the LD rule.
    /// </summary>
    public class CrtLayout12Record : StandardRecord
    {
        public const short sid = 0x89D;

        private short field_1_frtHeader_rt;
        private short field_2_frtHeader_grbitFrt;
        //private long field_3_frtHeader_reserved;
        private int field_5_dwCheckSum;
        private short field_6_option;
        private short field_7_wXMode;
        private short field_8_wYMode;
        private short field_9_wWidthMode;
        private short field_10_wHeightMode;

        private double field_11_x;
        private double field_12_y;
        private double field_13_dx;
        private double field_14_dy;

        public static int AutoLayoutType_Bottom = 0x0;
        public static int AutoLayoutType_TopRightCorner = 0x1;
        public static int AutoLayoutType_Top = 0x2;
        public static int AutoLayoutType_Right = 0x3;
        public static int AutoLayoutType_Left = 0x4;

        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CRTLAYOUT12]").AppendLine()
                .Append("   .rt               =").Append(HexDump.ToHex(field_1_frtHeader_rt)).Append("(").Append(field_1_frtHeader_rt).AppendLine(")")
                .Append("   .grbit            =").Append(HexDump.ToHex(field_2_frtHeader_grbitFrt)).Append("(").Append(field_2_frtHeader_grbitFrt).AppendLine(")")
                .Append("   .reserved         =").Append(HexDump.ToHex(0)).Append("(").Append(0).AppendLine(")")
                .Append("   .dwCheckSum       =").Append(HexDump.ToHex(field_5_dwCheckSum)).Append("(").Append(field_5_dwCheckSum).AppendLine(")")
                .Append("   .option           =").Append(HexDump.ToHex(field_6_option)).Append("(").Append(field_6_option).AppendLine(")")
                .Append("       .autolayouttype =").Append(autolayouttype.GetValue(field_6_option)).AppendLine()
                .Append("   .wXMode           =").Append(HexDump.ToHex(field_7_wXMode)).Append("(").Append(field_7_wXMode).AppendLine(")")
                .Append("   .wYMode           =").Append(HexDump.ToHex(field_8_wYMode)).Append("(").Append(field_8_wYMode).AppendLine(")")
                .Append("   .wWidthMode       =").Append(HexDump.ToHex(field_9_wWidthMode)).Append("(").Append(field_9_wWidthMode).AppendLine(")")
                .Append("   .wHeightMode      =").Append(HexDump.ToHex(field_10_wHeightMode)).Append("(").Append(field_10_wHeightMode).AppendLine(")")
                .Append("   .x                =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_11_x))).Append("(").Append(field_11_x).AppendLine(")")
                .Append("   .y                =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_12_y))).Append("(").Append(field_12_y).AppendLine(")")
                .Append("   .dx               =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_13_dx))).Append("(").Append(field_13_dx).AppendLine(")")
                .Append("   .dy               =").Append(HexDump.ToHex(BitConverter.DoubleToInt64Bits(field_14_dy))).Append("(").Append(field_14_dy).AppendLine(")")
                .AppendLine("[/CRTLAYOUT12]");

            return buffer.ToString();
        }
        BitField autolayouttype = BitFieldFactory.GetInstance(0x1E);

        //private short field_15_reserved2;

        public CrtLayout12Record()
        {
            field_1_frtHeader_rt = 0x89D;
            field_2_frtHeader_grbitFrt = 0;
        }

        public CrtLayout12Record(RecordInputStream ris)
        {
            field_1_frtHeader_rt = ris.ReadShort();
            field_2_frtHeader_grbitFrt = ris.ReadShort();
            ris.ReadLong();
            field_5_dwCheckSum = ris.ReadInt();
            field_6_option = ris.ReadShort();
            field_7_wXMode = ris.ReadShort();
            field_8_wYMode = ris.ReadShort();
            field_9_wWidthMode = ris.ReadShort();
            field_10_wHeightMode = ris.ReadShort();
            field_11_x = ris.ReadDouble();
            field_12_y = ris.ReadDouble();
            field_13_dx = ris.ReadDouble();
            field_14_dy = ris.ReadDouble();
            ris.ReadShort();
        }

        protected override int DataSize
        {
            get
            {
                return 2 + 2 + 4 + 4 +
                    4 +
                    2 +
                    2 + 2 + 2 + 2 +
                    8 + 8 + 8 + 8 + 2;
            }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_frtHeader_rt);
            out1.WriteShort(field_2_frtHeader_grbitFrt);
            out1.WriteInt(0);
            out1.WriteInt(0);
            out1.WriteInt(field_5_dwCheckSum);
            out1.WriteShort(field_6_option);
            out1.WriteShort(field_7_wXMode);
            out1.WriteShort(field_8_wYMode);
            out1.WriteShort(field_9_wWidthMode);
            out1.WriteShort(field_10_wHeightMode);
            out1.WriteDouble(field_11_x);
            out1.WriteDouble(field_12_y);
            out1.WriteDouble(field_13_dx);
            out1.WriteDouble(field_14_dy);
            out1.WriteShort(0);
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override object Clone()
        {
            CrtLayout12Record record = new CrtLayout12Record();
            record.AutoLayoutType = this.AutoLayoutType;
            record.CheckSum = this.CheckSum;
            record.DX = this.DX;
            record.DY = this.DY;
            record.HeightMode = this.HeightMode;
            record.WidthMode = this.WidthMode;
            record.X = this.X;
            record.XMode = this.XMode;
            record.Y = this.Y;
            record.YMode = this.YMode;
            return record;
        }
        
        /// <summary>
        /// automatic layout type of the legend. 
        /// MUST be ignored when this record is in the sequence of records that conforms to the ATTACHEDLABEL rule. 
        /// MUST be a value from the following table:
        /// 0x0  Align to the bottom
        /// 0x1  Align to top right corner
        /// 0x2  Align to the top
        /// 0x3  Align to the right
        /// 0x4  Align to the left
        /// </summary>
        public int AutoLayoutType
        {
            get { return autolayouttype.GetValue(field_6_option); }
            set { field_6_option = autolayouttype.SetShortValue(field_6_option, (short)value); }
        }
        /// <summary>
        /// specifies the checksum of the values in the order as follows,
        /// </summary>
        public int CheckSum
        {
            get { return field_5_dwCheckSum; }
            set { field_5_dwCheckSum = value; }
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

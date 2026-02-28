using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// The Chart3d record specifies that the plot area of the chart group is rendered in a 3-D scene 
    /// and also specifies the attributes of the 3-D plot area. The preceding chart group type MUST be 
    /// of type bar, pie, line, area, or surface.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class Chart3dRecord: StandardRecord
    {
        public const short sid = 0x103A;


        private short field_1_anRot;
        private short field_2_anElev;
        private short field_3_pcDist;
        private short field_4_pcHeight;
        private short field_5_pcDepth;
        private short field_6_pcGap;
        private short field_7_option;
        private BitField fPerspective = BitFieldFactory.GetInstance(0x1);
        private BitField fCluster = BitFieldFactory.GetInstance(0x2);
        private BitField f3DScaling = BitFieldFactory.GetInstance(0x4);
        private BitField reserved1 = BitFieldFactory.GetInstance(0x8);
        private BitField fNotPieChart = BitFieldFactory.GetInstance(0x10);
        private BitField fWalls2D = BitFieldFactory.GetInstance(0x20);


        public Chart3dRecord()
        {
        }

        public Chart3dRecord(RecordInputStream in1)
        {
            field_1_anRot = in1.ReadShort();
            field_2_anElev = in1.ReadShort();
            field_3_pcDist = in1.ReadShort();
            field_4_pcHeight = in1.ReadShort();
            field_5_pcDepth = in1.ReadShort();
            field_6_pcGap = in1.ReadShort();
            field_7_option = in1.ReadShort();
        }

        protected override int DataSize
        {
            get { return 14; }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_anRot);
            out1.WriteShort(field_2_anElev);
            out1.WriteShort(field_3_pcDist);
            out1.WriteShort(field_4_pcHeight);
            out1.WriteShort(field_5_pcDepth);
            out1.WriteShort(field_6_pcGap);
            out1.WriteShort(field_7_option);
        }

        public override short Sid
        {
            get { return sid; }
        }
        /// <summary>
        /// A signed integer that specifies the clockwise rotation, in degrees, of the 3-D plot area 
        /// around a vertical line through the center of the 3-D plot area. MUST be greater than or 
        /// equal to 0 and MUST be less than or equal to 360.
        /// </summary>
        public int Rotation
        {
            get { return field_1_anRot; }
            set 
            {
                if (value < 0) value = 0;
                if (value > 360) value = 360;
                field_1_anRot = (short)value; 
            }
        }

        /// <summary>
        /// A signed integer that specifies the rotation, in degrees, of the 3-D plot area around 
        /// a horizontal line through the center of the 3-D plot area.MUST be greater than or equal 
        /// to -90 and MUST be less than or equal to 90.
        /// </summary>
        public int Elev
        {
            get { return field_2_anElev; }
            set
            {
                if (value < -90) value = -90;
                if (value > 90) value = 90;
                field_2_anElev = (short)value;
            }
        }
        /// <summary>
        /// A signed integer that specifies the field of view angle for the 3-D plot area. 
        /// MUST be greater than or equal to zero and less than 200.
        /// </summary>
        public int Dist
        {
            get { return field_3_pcDist; }
            set
            {
                if (value < 0) value = 0;
                if (value > 200) value = 200;
                field_3_pcDist = (short)value;
            }
        }
        /// <summary>
        /// If fNotPieChart is 0, then this is an unsigned integer that specifies the thickness of the pie for a pie chart group. 
        /// If fNotPieChart is 1, then this is a signed integer that specifies the height of the 3-D plot area as a percentage of its width.
        /// </summary>
        public int Height
        {
            get { return field_4_pcHeight; }
            set {
                field_4_pcHeight = (short)value;
            }
        }

        /// <summary>
        /// A signed integer that specifies the depth of the 3-D plot area as a percentage of its width.
        /// MUST be greater than or equal to 1 and less than or equal to 2000.
        /// </summary>
        public int Depth
        {
            get { return field_5_pcDepth; }
            set { field_5_pcDepth = (short)value; }
        }

        /// <summary>
        /// An unsigned integer that specifies the width of the gap between the series and the front and 
        /// back edges of the 3-D plot area as a percentage of the data point depth divided by 2. 
        /// If fCluster is not 1 and chart group type is not a bar then pcGap also specifies distance 
        /// between adjacent series as a percentage of the data point depth. MUST be less than or equal to 500.
        /// </summary>
        public int Gap
        {
            get { return field_6_pcGap; }
            set { field_6_pcGap = (short)value; }
        }

        /// <summary>
        /// A bit that specifies whether the 3-D plot area is rendered with a vanishing point. 
        /// If fNotPieChart is 0 the value MUST be 0. If fNotPieChart is 1 then the value 
        /// MUST be a value from the following 
        /// true   Perspective vanishing point applied based on value of pcDist.
        /// false  No vanishing point applied.
        /// </summary>
        public bool IsPerspective
        {
            get { return fPerspective.IsSet(field_7_option); }
            set { field_7_option = fPerspective.SetShortBoolean(field_7_option, value); }
        }

        /// <summary>
        /// specifies whether data points are clustered together in a bar chart group. 
        /// If chart group type is not bar or pie, value MUST be ignored. If chart group type is pie,
        /// value MUST be 0. If chart group type is bar, then the value MUST be a value from the following
        /// true   Data points are clustered.
        /// false  Data points are not clustered.
        /// </summary>
        public bool IsCluster
        {
            get { return fCluster.IsSet(field_7_option); }
            set { field_7_option = fCluster.SetShortBoolean(field_7_option, value); }
        }

        /// <summary>
        /// A bit that specifies whether the height of the 3-D plot area is automatically determined. 
        /// If fNotPieChart is 0 then this MUST be 0. If fNotPieChart is 1 then the value MUST be a value from the following table:
        /// false The value of pcHeight is used to determine the height of the 3-D plot area
        /// true  The height of the 3-D plot area is automatically determined
        /// </summary>
        public bool Is3DScaling
        {
            get { return f3DScaling.IsSet(field_7_option); }
            set { field_7_option = f3DScaling.SetShortBoolean(field_7_option, value); }
        }
        /// <summary>
        /// A bit that specifies whether the chart group type is pie. MUST be a value from the following :
        /// false Chart group type MUST be pie.
        /// true  Chart group type MUST not be pie.
        /// </summary>
        public bool IsNotPieChart
        {
            get { return fNotPieChart.IsSet(field_7_option); }
            set { field_7_option = fNotPieChart.SetShortBoolean(field_7_option, value); }
        }

        /// <summary>
        /// Whether the walls are rendered in 2-D. If fPerspective is 1 then this MUST be ignored. 
        /// If the chart group type is not bar, area or pie this MUST be ignored. 
        /// If the chart group is of type bar and fCluster is 0, then this MUST be ignored. 
        /// If the chart group type is pie this MUST be 0 and MUST be ignored. 
        /// If the chart group type is bar or area, then the value MUST be a value from the following
        /// false  Chart walls and floor are rendered in 3D.
        /// true   Chart walls are rendered in 2D and the chart floor is not rendered.
        /// </summary>
        public bool IsWalls2D
        {
            get { return fWalls2D.IsSet(field_7_option); }
            set { field_7_option = fWalls2D.SetShortBoolean(field_7_option, value); }
        }

        public override object Clone()
        {
            Chart3dRecord record = new Chart3dRecord();
            record.Depth = this.Depth;
            record.Dist = this.Dist;
            record.Elev = this.Elev;
            record.Height = this.Height;
            record.Gap = this.Gap;
            record.Is3DScaling = this.Is3DScaling;
            record.IsCluster = this.IsCluster;
            record.IsNotPieChart = this.IsNotPieChart;
            record.IsPerspective = this.IsPerspective;
            record.IsWalls2D = this.IsWalls2D;
            record.Rotation = this.Rotation;

            return record;
        }
        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CHART3D]").AppendLine()
                .Append("   .anRot              =").Append(HexDump.ToHex(field_1_anRot)).Append("(").Append(field_1_anRot).AppendLine(")")
                .Append("   .anElev             =").Append(HexDump.ToHex(field_2_anElev)).Append("(").Append(field_2_anElev).AppendLine(")")
                .Append("   .pcDist             =").Append(HexDump.ToHex(field_3_pcDist)).Append("(").Append(field_3_pcDist).AppendLine(")")
                .Append("   .pcHeight           =").Append(HexDump.ToHex(field_4_pcHeight)).Append("(").Append(field_4_pcHeight).AppendLine(")")
                .Append("   .pcDepth            =").Append(HexDump.ToHex(field_5_pcDepth)).Append("(").Append(field_5_pcDepth).AppendLine(")")
                .Append("   .pcGap              =").Append(HexDump.ToHex(field_6_pcGap)).Append("(").Append(field_6_pcGap).AppendLine(")")
                .Append("   .option             =").Append(HexDump.ToHex(field_7_option)).Append("(").Append(field_7_option).AppendLine(")")
                .Append("       .fPerspective       =").Append(IsPerspective).AppendLine()
                .Append("       .fCluster           =").Append(IsCluster).AppendLine()
                .Append("       .f3DScaling         =").Append(Is3DScaling).AppendLine()
                .Append("       .fNotPieChart       =").Append(IsNotPieChart).AppendLine()
                .Append("       .fWalls2D           =").Append(IsWalls2D).AppendLine()
                .AppendLine("[/CHART3D]");

            return buffer.ToString();
        }
    }
}

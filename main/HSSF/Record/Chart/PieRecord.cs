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
    /// The Pie record specifies that the chart group is a pie chart group or 
    /// a doughnut chart group, and specifies the chart group attributes.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class PieRecord : StandardRecord
    {
        public const short sid = 0x1019;

        private short field_1_anStart;
        private short field_2_pcDonut;
        private short field_3_option;

        private BitField fHasShadow = BitFieldFactory.GetInstance(0x1);
        private BitField fShowLdrLines = BitFieldFactory.GetInstance(0x2);

        public PieRecord()
        {
        }

        public PieRecord(RecordInputStream ris)
        {
            field_1_anStart = ris.ReadShort();
            field_2_pcDonut = ris.ReadShort();
            field_3_option = ris.ReadShort();
        }


        protected override int DataSize
        {
            get { return 2 + 2 + 2; }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_anStart);
            out1.WriteShort(field_2_pcDonut);
            out1.WriteShort(field_3_option);
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override object Clone()
        {
            PieRecord record = new PieRecord();
            record.Dount = this.Dount;
            record.HasShadow = this.HasShadow;
            record.ShowLdrLines = this.ShowLdrLines;
            record.Start = this.Start;
            return record;
        }

        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PIE]").AppendLine()
                .Append("   .anStart             =").Append(HexDump.ToHex(field_1_anStart)).Append("(").Append(field_1_anStart).AppendLine(")")
                .Append("   .anDount             =").Append(HexDump.ToHex(field_2_pcDonut)).Append("(").Append(field_2_pcDonut).AppendLine(")")
                .Append("   .option              =").Append(HexDump.ToHex(field_3_option)).Append("(").Append(field_3_option).AppendLine(")")
                .Append("       .fHasShadow         =").Append(HasShadow).AppendLine()
                .Append("       .fShowLdrLines      =").Append(ShowLdrLines).AppendLine()
                .AppendLine("[/PIE]");

            return buffer.ToString();
        }

        /// <summary>
        /// An unsigned integer that specifies the starting angle of the first data point, 
        /// clockwise from the top of the circle. MUST be less than or equal to 360.
        /// </summary>
        public int Start
        {
            get { return field_1_anStart; }
            set { field_1_anStart = (short)value; }
        }

        /// <summary>
        /// An unsigned integer that specifies the size of the center hole in a doughnut chart group 
        /// as a percentage of the plot area size. MUST be a value from the following table:
        /// 0          Pie chart group.
        /// 10 to 90   Doughnut chart group.
        /// </summary>
        public int Dount
        {
            get { return field_2_pcDonut; }
            set { field_2_pcDonut = (short)value; }
        }

        /// <summary>
        /// A bit that specifies whether one data point or more data points in the chart group have shadows.
        /// </summary>
        public bool HasShadow
        {
            get { return fHasShadow.IsSet(field_3_option); }
            set { field_3_option = fHasShadow.SetShortBoolean(field_3_option, value); }
        }

        /// <summary>
        /// A bit that specifies whether the leader lines to the data labels are shown.
        /// </summary>
        public bool ShowLdrLines
        {
            get { return fShowLdrLines.IsSet(field_3_option); }
            set { field_3_option = fShowLdrLines.SetShortBoolean(field_3_option, value); }
        }
    }
}

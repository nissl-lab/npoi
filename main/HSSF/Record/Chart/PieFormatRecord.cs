using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// The PieFormat record specifies the distance of a data point or data points in a series from the center of one of the following:
    /// The plot area for a doughnut or pie chart group.
    /// The primary pie in a pie of pie or bar of pie chart group.
    /// The secondary bar/pie of a pie of pie chart group.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class PieFormatRecord : StandardRecord
    {
        public const short sid = 0x100B;

        private short field_1_pcExplode;

        public PieFormatRecord()
        {
        }

        public PieFormatRecord(RecordInputStream ris)
        {
            field_1_pcExplode = ris.ReadShort();
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_pcExplode);
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override object Clone()
        {
            PieFormatRecord record = new PieFormatRecord();

            record.Explode = this.Explode;

            return record;
        }

        public override string ToString()
        {

            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PIEFORMAT]").AppendLine()
                .Append("   .pcExplode = ").Append(HexDump.ToHex(field_1_pcExplode)).Append("(").Append(field_1_pcExplode).AppendLine(")")
                .AppendLine("[/PIEFORMAT]");

            return buffer.ToString();
        }
        /// <summary>
        /// A signed integer that specifies the distance of a data point or data points in a series from the center of one of the following:
        /// The plot area for a doughnut or pie chart group.
        /// The primary pie in a pie of pie or bar of pie chart group.
        /// The secondary bar/pie of a pie of pie chart group.
        /// </summary>
        public int Explode
        {
            get { return field_1_pcExplode; }
            set { field_1_pcExplode = (short)value; }
        }
    }
}

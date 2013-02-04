
namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// The Fbi2 record specifies the font information at the time the scalable font is added to the chart.
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    public class Fbi2Record : RowDataRecord
    {
        public const short sid = 0x1068;
        public Fbi2Record(RecordInputStream ris)
            : base(ris)
        {
        }

        protected override int DataSize
        {
            get
            {
                return base.DataSize;
            }
        }
        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            base.Serialize(out1);
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}

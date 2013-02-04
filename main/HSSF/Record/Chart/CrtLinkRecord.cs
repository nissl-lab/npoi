
namespace NPOI.HSSF.Record.Chart
{
    /// <summary>
    /// The CrtLink record is written but unused.
    /// </summary>
    public class CrtLinkRecord:StandardRecord
    {
        //0x1022
        public const short sid = 0x1022;

        public CrtLinkRecord()
        { 
        }

        public CrtLinkRecord(RecordInputStream in1)
        {
            in1.ReadInt();
            in1.ReadInt();
            in1.ReadShort();
        }

        protected override int DataSize
        {
            get { return 10; }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteInt(0);
            out1.WriteInt(0);
            out1.WriteShort(0);
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override object Clone()
        {
            return new CrtLinkRecord();
        }
        public override string ToString()
        {
            return "[CrtLink]Unused[/CrtLink]";
        }
    }
}

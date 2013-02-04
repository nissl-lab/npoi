
namespace NPOI.HSSF.Record.AutoFilter
{
    public class FilterModeRecord:StandardRecord
    {
        public FilterModeRecord()
        { 
        }

        public FilterModeRecord(RecordInputStream in1)
        { 
        
        }

        public const short sid = 0x9b;
        public override short Sid
        {
            get { return sid; }
        }

        protected override int DataSize
        {
            get { return 0; }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            
        }
        public override object Clone()
        {
            FilterModeRecord rec = new FilterModeRecord();
            return rec;
        }
    }
}

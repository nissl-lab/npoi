using System;
using System.Text;

namespace NPOI.HSSF.Record.AutoFilter
{
    public class AutoFilterInfoRecord:StandardRecord
    {

        public AutoFilterInfoRecord()
        { 
        
        }

        private short field_1_cEntries = 0;

        public AutoFilterInfoRecord(RecordInputStream in1)
        {
            field_1_cEntries = in1.ReadShort();
        }

        public const short sid = 0x9D;
        public override short Sid
        {
            get { return sid; }
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public short NumEntries
        {
            get { return field_1_cEntries; }
            set { field_1_cEntries = value; }
        }
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[AUTOFILTERINFO]\n");
            buffer.Append("    .numEntries          = ")
                .Append(field_1_cEntries).Append("\n");
            buffer.Append("[/AUTOFILTERINFO]\n");
            return buffer.ToString();
        }
        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_cEntries);
        }
        public override object Clone()
        {
            AutoFilterInfoRecord rec = new AutoFilterInfoRecord();
            rec.field_1_cEntries = field_1_cEntries;
            return rec;
        }
    }
}

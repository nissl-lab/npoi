using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF.Record
{
    public class Excel9FileRecord:StandardRecord
    {
        public Excel9FileRecord()
        { 
        
        }
        public Excel9FileRecord(RecordInputStream in1)
        { 
        
        }

        public static short sid = 0x1c0;
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
    }
}

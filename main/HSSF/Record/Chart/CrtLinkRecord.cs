using System;
using System.Collections.Generic;
using System.Text;

using NPOI.Util;

namespace NPOI.HSSF.Record.Chart
{
    [Obsolete("This class was not found in poi")]
    public class CrtLinkRecord:StandardRecord
    {
        public const short sid = 4130;

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

        public override string ToString()
        {
            return "[CrtLink]Unused[/CrtLink]";
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;
    using NPOI.Util.IO;

    public class PrintSizeRecord:StandardRecord
    {
        public static short sid = 0x33;

        private short printSize;

        public PrintSizeRecord(RecordInputStream in1)
        {
            printSize = in1.ReadShort();
        }

        public short PrintSize
        {
            get { return printSize; }
            set { printSize = value; }
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override void Serialize(LittleEndianOutput out1)
        {
            out1.WriteShort(printSize);
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}

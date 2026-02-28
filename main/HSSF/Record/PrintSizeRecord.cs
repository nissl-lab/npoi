
namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;

    public class PrintSizeRecord:StandardRecord
    {
        public const short sid = 0x33;

        private short printSize;

        public PrintSizeRecord()
        { 
        
        }

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

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(printSize);
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override object Clone()
        {
            PrintSizeRecord pzr = new PrintSizeRecord();
            pzr.printSize = this.printSize;
            return pzr;
        }
    }
}

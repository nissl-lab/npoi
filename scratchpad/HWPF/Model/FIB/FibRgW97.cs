using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FibRgW97
    {
        private short lidFE;

        public FibRgW97()
        {
        }

        public void Deserialize(HWPFStream stream)
        {
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            stream.ReadShort();
            lidFE = stream.ReadShort();            
        }

        public short LIDOfStoredStyleName
        {
            get { return lidFE; }
            set { lidFE = value; }
        }


    }
}

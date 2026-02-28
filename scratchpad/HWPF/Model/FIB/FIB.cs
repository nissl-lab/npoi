using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FIB
    {
        public FIBBase fibBase;
        public short csw = 0;
        public FibRgW97 fibRgW97;
        public short cslw = 0;
        public FibRgLw97 fibRgLw97;

        public FibRgFcLcb fibRgFcLcb;
        public short cswNew = 0;
        public FIBCswNew fibCswNew;

        public FIB()
        {
            
        }

        public void Deserialize(HWPFStream stream)
        {
            fibBase = new FIBBase();
            fibBase.Deserialize(stream);
            csw = stream.ReadShort();
            if (csw != 0x000E)
            {
                throw new ArgumentOutOfRangeException("csw must be 0x000E");
            }
            fibRgW97 = new FibRgW97();
            fibRgW97.Deserialize(stream);
            cslw = stream.ReadShort();
            if (cslw != 0x0016)
            {
                throw new ArgumentOutOfRangeException("cslw must be 0x0016");
            }
            fibRgLw97 = new FibRgLw97();
            fibRgLw97.Deserialize(stream);
            
            fibRgFcLcb = new FibRgFcLcb();
            fibRgFcLcb.Deserialize(stream);
            cswNew = stream.ReadShort();
            fibCswNew = new FIBCswNew();
            fibCswNew.Deserialize(stream);
        }
    }
}

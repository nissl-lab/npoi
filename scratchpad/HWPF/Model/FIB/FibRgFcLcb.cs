using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FibRgFcLcb
    {
        public FibRgFcLcb97 fibRgFcLcb97;
        public FibRgFcLcb2000 fibRgFcLcb2000;
        public FibRgFcLcb2002 fibRgFcLcb2002;
        public FibRgFcLcb2003 fibRgFcLcb2003;
        public FibRgFcLcb2007 fibRgFcLcb2007;
        public short cbRgFcLcb = 0;

        public FibRgFcLcb()
        {
            
        }

        public void Deserialize(HWPFStream stream)
        {
            cbRgFcLcb = stream.ReadShort();
            switch (cbRgFcLcb)
            {
                case 0x005D:
                    fibRgFcLcb97 = new FibRgFcLcb97();
                    fibRgFcLcb97.Deserialize(stream);
                    break;
                case 0x006C:
                    fibRgFcLcb2000 = new FibRgFcLcb2000();
                    fibRgFcLcb2000.Deserialize(stream);
                    fibRgFcLcb97 = fibRgFcLcb2000.fibRgFcLcb97;
                    break;
                case 0x0088:
                    fibRgFcLcb2002 = new FibRgFcLcb2002();
                    fibRgFcLcb2002.Deserialize(stream);
                    fibRgFcLcb97 = fibRgFcLcb2002.fibRgFcLcb2000.fibRgFcLcb97;
                    fibRgFcLcb2000 = fibRgFcLcb2002.fibRgFcLcb2000;
                    break;
                case 0x00A4 :
                    fibRgFcLcb2003 = new FibRgFcLcb2003();
                    fibRgFcLcb2003.Deserialize(stream);
                    fibRgFcLcb97 = fibRgFcLcb2003.fibRgFcLcb2002.fibRgFcLcb2000.fibRgFcLcb97;
                    fibRgFcLcb2000 = fibRgFcLcb2003.fibRgFcLcb2002.fibRgFcLcb2000;
                    fibRgFcLcb2002 = fibRgFcLcb2003.fibRgFcLcb2002;
                    break;
                case 0x00B7:
                    fibRgFcLcb2007 = new FibRgFcLcb2007();
                    fibRgFcLcb2007.Deserialize(stream);
                    fibRgFcLcb97 = fibRgFcLcb2007.fibRgFcLcb2003.fibRgFcLcb2002.fibRgFcLcb2000.fibRgFcLcb97;
                    fibRgFcLcb2000 = fibRgFcLcb2007.fibRgFcLcb2003.fibRgFcLcb2002.fibRgFcLcb2000;
                    fibRgFcLcb2002 = fibRgFcLcb2007.fibRgFcLcb2003.fibRgFcLcb2002;
                    fibRgFcLcb2003 = fibRgFcLcb2007.fibRgFcLcb2003;
                    break;
            }
        }
    }
}

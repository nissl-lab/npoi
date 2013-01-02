using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FibRgFcLcb2000
    {
        public int fcPlcfTch = 0;
        public int lcbPlcfTch = 0;
        public int fcRmdThreading = 0;
        public int lcbRmdThreading = 0;
        public int fcMid = 0;
        public int lcbMid = 0;
        public int fcSttbRgtplc = 0;
        public int lcbSttbRgtplc = 0;
        public int fcMsoEnvelope = 0;
        public int lcbMsoEnvelope = 0;
        public int fcPlcfLad = 0;
        public int lcbPlcfLad = 0;
        public int fcRgDofr = 0;
        public int lcbRgDofr = 0;
        public int fcPlcosl = 0;
        public int lcbPlcosl = 0;
        public int fcPlcfCookieOld = 0;
        public int lcbPlcfCookieOld = 0;
        public int fcPgdMotherOld = 0;
        public int lcbPgdMotherOld = 0;
        public int fcBkdMotherOld = 0;
        public int lcbBkdMotherOld = 0;
        public int fcPgdFtnOld = 0;
        public int lcbPgdFtnOld = 0;
        public int fcBkdFtnOld = 0;
        public int lcbBkdFtnOld = 0;
        public int fcPgdEdnOld = 0;
        public int lcbPgdEdnOld = 0;
        public int fcBkdEdnOld = 0;
        public int lcbBkdEdnOld = 0;

        public FibRgFcLcb97 fibRgFcLcb97;

        public FibRgFcLcb2000()
        {
        }

        public void Deserialize(HWPFStream stream)
        {
            fibRgFcLcb97 = new FibRgFcLcb97();
            fibRgFcLcb97.Deserialize(stream);

            fcPlcfTch = stream.ReadInt();
            lcbPlcfTch = stream.ReadInt();
            fcRmdThreading = stream.ReadInt();
            lcbRmdThreading = stream.ReadInt();
            fcMid = stream.ReadInt();
            lcbMid = stream.ReadInt();
            fcSttbRgtplc = stream.ReadInt();
            lcbSttbRgtplc = stream.ReadInt();
            fcMsoEnvelope = stream.ReadInt();
            lcbMsoEnvelope = stream.ReadInt();
            fcPlcfLad = stream.ReadInt();
            lcbPlcfLad = stream.ReadInt();
            fcRgDofr = stream.ReadInt();
            lcbRgDofr = stream.ReadInt();
            fcPlcosl = stream.ReadInt();
            lcbPlcosl = stream.ReadInt();
            fcPlcfCookieOld = stream.ReadInt();
            lcbPlcfCookieOld = stream.ReadInt();
            fcPgdMotherOld = stream.ReadInt();
            lcbPgdMotherOld = stream.ReadInt();
            fcBkdMotherOld = stream.ReadInt();
            lcbBkdMotherOld = stream.ReadInt();
            fcPgdFtnOld = stream.ReadInt();
            lcbPgdFtnOld = stream.ReadInt();
            fcBkdFtnOld = stream.ReadInt();
            lcbBkdFtnOld = stream.ReadInt();
            fcPgdEdnOld = stream.ReadInt();
            lcbPgdEdnOld = stream.ReadInt();
            fcBkdEdnOld = stream.ReadInt();
            lcbBkdEdnOld = stream.ReadInt();
            
        }
    }
}

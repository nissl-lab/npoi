using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FibRgLw97
    {
        int cbMac=0;
        int ccpText=0;
        int ccpFtn=0;
        int ccpHdd=0;
        int ccpAtn=0;
        int ccpEdn = 0;
        int ccpTxbx = 0;
        int ccpHdrTxbx = 0;

        public FibRgLw97()
        {

        }

        public void Deserialize(HWPFStream stream)
        {
            cbMac = stream.ReadInt();
            stream.ReadInt();
            stream.ReadInt();
            ccpText = stream.ReadInt();
            ccpFtn = stream.ReadInt();
            ccpHdd = stream.ReadInt();
            stream.ReadInt();
            ccpAtn = stream.ReadInt();
            ccpEdn = stream.ReadInt();
            ccpTxbx = stream.ReadInt();
            ccpHdrTxbx = stream.ReadInt();

            stream.ReadInt(); //reserved4
            stream.ReadInt(); //reserved5
            stream.ReadInt();//reserved6
            stream.ReadInt();//reserved7
            stream.ReadInt();//reserved8
            stream.ReadInt();//reserved9
            stream.ReadInt();//reserved10
            stream.ReadInt();//reserved11
            stream.ReadInt();//reserved12
            stream.ReadInt();//reserved13
            stream.ReadInt();//reserved14           
        }
    }
}

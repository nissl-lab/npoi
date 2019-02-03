using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.OpenXmlFormats.Shared
{
    public interface IOMathContainer
    {
        ArrayList Items { get; }

        CT_R AddNewR();
        CT_Acc AddNewAcc();
        CT_Nary AddNewNary();
        CT_SSub AddNewSSub();
        CT_F AddNewF();
    }
}

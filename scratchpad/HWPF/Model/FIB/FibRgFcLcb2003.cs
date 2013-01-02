using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FibRgFcLcb2003
    {
        public int fcHplxsdr= 0;
        public int lcbHplxsdr= 0;
        public int fcSttbfBkmkSdt= 0;
        public int lcbSttbfBkmkSdt= 0;
        public int fcPlcfBkfSdt= 0;
        public int lcbPlcfBkfSdt= 0;
        public int fcPlcfBklSdt= 0;
        public int lcbPlcfBklSdt= 0;
        public int fcCustomXForm= 0;
        public int lcbCustomXForm= 0;
        public int fcSttbfBkmkProt= 0;
        public int lcbSttbfBkmkProt= 0;
        public int fcPlcfBkfProt= 0;
        public int lcbPlcfBkfProt= 0;
        public int fcPlcfBklProt= 0;
        public int lcbPlcfBklProt= 0;
        public int fcSttbProtUser= 0;
        public int lcbSttbProtUser= 0;
        public int fcUnused= 0;
        public int lcbUnused= 0;
        public int fcPlcfpmiOld= 0;
        public int lcbPlcfpmiOld= 0;
        public int fcPlcfpmiOldInline= 0;
        public int lcbPlcfpmiOldInline= 0;
        public int fcPlcfpmiNew= 0;
        public int lcbPlcfpmiNew= 0;
        public int fcPlcfpmiNewInline= 0;
        public int lcbPlcfpmiNewInline= 0;
        public int fcPlcflvcOld= 0;
        public int lcbPlcflvcOld= 0;
        public int fcPlcflvcOldInline= 0;
        public int lcbPlcflvcOldInline= 0;
        public int fcPlcflvcNew= 0;
        public int lcbPlcflvcNew= 0;
        public int fcPlcflvcNewInline= 0;
        public int lcbPlcflvcNewInline= 0;
        public int fcPgdMother= 0;
        public int lcbPgdMother= 0;
        public int fcBkdMother= 0;
        public int lcbBkdMother= 0;
        public int fcAfdMother= 0;
        public int lcbAfdMother= 0;
        public int fcPgdFtn= 0;
        public int lcbPgdFtn= 0;
        public int fcBkdFtn= 0;
        public int lcbBkdFtn= 0;
        public int fcAfdFtn= 0;
        public int lcbAfdFtn= 0;
        public int fcPgdEdn= 0;
        public int lcbPgdEdn= 0;
        public int fcBkdEdn= 0;
        public int lcbBkdEdn= 0;
        public int fcAfdEdn= 0;
        public int lcbAfdEdn= 0;
        public int fcAfd= 0;
        public int lcbAfd= 0;

        public FibRgFcLcb2002 fibRgFcLcb2002;

        public FibRgFcLcb2003()
        {

        }

        public void Deserialize(HWPFStream stream)
        {
            fibRgFcLcb2002 = new FibRgFcLcb2002();
            fibRgFcLcb2002.Deserialize(stream);

            fcHplxsdr = stream.ReadInt();
            lcbHplxsdr = stream.ReadInt();
            fcSttbfBkmkSdt = stream.ReadInt();
            lcbSttbfBkmkSdt = stream.ReadInt();
            fcPlcfBkfSdt = stream.ReadInt();
            lcbPlcfBkfSdt = stream.ReadInt();
            fcPlcfBklSdt = stream.ReadInt();
            lcbPlcfBklSdt = stream.ReadInt();
            fcCustomXForm = stream.ReadInt();
            lcbCustomXForm = stream.ReadInt();
            fcSttbfBkmkProt = stream.ReadInt();
            lcbSttbfBkmkProt = stream.ReadInt();
            fcPlcfBkfProt = stream.ReadInt();
            lcbPlcfBkfProt = stream.ReadInt();
            fcPlcfBklProt = stream.ReadInt();
            lcbPlcfBklProt = stream.ReadInt();
            fcSttbProtUser = stream.ReadInt();
            lcbSttbProtUser = stream.ReadInt();
            fcUnused = stream.ReadInt();
            lcbUnused = stream.ReadInt();
            fcPlcfpmiOld = stream.ReadInt();
            lcbPlcfpmiOld = stream.ReadInt();
            fcPlcfpmiOldInline = stream.ReadInt();
            lcbPlcfpmiOldInline = stream.ReadInt();
            fcPlcfpmiNew = stream.ReadInt();
            lcbPlcfpmiNew = stream.ReadInt();
            fcPlcfpmiNewInline = stream.ReadInt();
            lcbPlcfpmiNewInline = stream.ReadInt();
            fcPlcflvcOld = stream.ReadInt();
            lcbPlcflvcOld = stream.ReadInt();
            fcPlcflvcOldInline = stream.ReadInt();
            lcbPlcflvcOldInline = stream.ReadInt();
            fcPlcflvcNew = stream.ReadInt();
            lcbPlcflvcNew = stream.ReadInt();
            fcPlcflvcNewInline = stream.ReadInt();
            lcbPlcflvcNewInline = stream.ReadInt();
            fcPgdMother = stream.ReadInt();
            lcbPgdMother = stream.ReadInt();
            fcBkdMother = stream.ReadInt();
            lcbBkdMother = stream.ReadInt();
            fcAfdMother = stream.ReadInt();
            lcbAfdMother = stream.ReadInt();
            fcPgdFtn = stream.ReadInt();
            lcbPgdFtn = stream.ReadInt();
            fcBkdFtn = stream.ReadInt();
            lcbBkdFtn = stream.ReadInt();
            fcAfdFtn = stream.ReadInt();
            lcbAfdFtn = stream.ReadInt();
            fcPgdEdn = stream.ReadInt();
            lcbPgdEdn = stream.ReadInt();
            fcBkdEdn = stream.ReadInt();
            lcbBkdEdn = stream.ReadInt();
            fcAfdEdn = stream.ReadInt();
            lcbAfdEdn = stream.ReadInt();
            fcAfd = stream.ReadInt();
            lcbAfd = stream.ReadInt();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FibRgFcLcb2002
    {
        public int fcUnused1 = 0;
        public int lcbUnused1 = 0;
        public int fcPlcfPgp = 0;
        public int lcbPlcfPgp = 0;
        public int fcPlcfuim = 0;
        public int lcbPlcfuim = 0;
        public int fcPlfguidUim = 0;
        public int lcbPlfguidUim = 0;
        public int fcAtrdExtra = 0;
        public int lcbAtrdExtra = 0;
        public int fcPlrsid = 0;
        public int lcbPlrsid = 0;
        public int fcSttbfBkmkFactoid = 0;
        public int lcbSttbfBkmkFactoid = 0;
        public int fcPlcfBkfFactoid = 0;
        public int lcbPlcfBkfFactoid = 0;
        public int fcPlcpublicfcookie = 0;
        public int lcbPlcpublicfcookie = 0;
        public int fcPlcfBklFactoid = 0;
        public int lcbPlcfBklFactoid = 0;
        public int fcFactoidData = 0;
        public int lcbFactoidData = 0;
        public int fcDocUndo = 0;
        public int lcbDocUndo = 0;
        public int fcSttbfBkmkFcc = 0;
        public int lcbSttbfBkmkFcc = 0;
        public int fcPlcfBkfFcc = 0;
        public int lcbPlcfBkfFcc = 0;
        public int fcPlcfBklFcc = 0;
        public int lcbPlcfBklFcc = 0;
        public int fcSttbfbkmkBPRepairs = 0;
        public int lcbSttbfbkmkBPRepairs = 0;
        public int fcPlcfbkfBPRepairs = 0;
        public int lcbPlcfbkfBPRepairs = 0;
        public int fcPlcfbklBPRepairs = 0;
        public int lcbPlcfbklBPRepairs = 0;
        public int fcPmsNew = 0;
        public int lcbPmsNew = 0;
        public int fcODSO = 0;
        public int lcbODSO = 0;
        public int fcPlcfpmiOldXP = 0;
        public int lcbPlcfpmiOldXP = 0;
        public int fcPlcfpmiNewXP = 0;
        public int lcbPlcfpmiNewXP = 0;
        public int fcPlcfpmiMixedXP = 0;
        public int lcbPlcfpmiMixedXP = 0;
        public int fcUnused2 = 0;
        public int lcbUnused2 = 0;
        public int fcPlcffactoid = 0;
        public int lcbPlcffactoid = 0;
        public int fcPlcflvcOldXP = 0;
        public int lcbPlcflvcOldXP = 0;
        public int fcPlcflvcNewXP = 0;
        public int lcbPlcflvcNewXP = 0;
        public int fcPlcflvcMixedXP = 0;
        public int lcbPlcflvcMixedXP = 0;

        public FibRgFcLcb2000 fibRgFcLcb2000;

        public FibRgFcLcb2002()
        {
          

        }
        public void Deserialize(HWPFStream stream)
        {
            fibRgFcLcb2000 = new FibRgFcLcb2000();
            fibRgFcLcb2000.Deserialize(stream);

            fcUnused1 = stream.ReadInt();
            lcbUnused1 = stream.ReadInt();
            fcPlcfPgp = stream.ReadInt();
            lcbPlcfPgp = stream.ReadInt();
            fcPlcfuim = stream.ReadInt();
            lcbPlcfuim = stream.ReadInt();
            fcPlfguidUim = stream.ReadInt();
            lcbPlfguidUim = stream.ReadInt();
            fcAtrdExtra = stream.ReadInt();
            lcbAtrdExtra = stream.ReadInt();
            fcPlrsid = stream.ReadInt();
            lcbPlrsid = stream.ReadInt();
            fcSttbfBkmkFactoid = stream.ReadInt();
            lcbSttbfBkmkFactoid = stream.ReadInt();
            fcPlcfBkfFactoid = stream.ReadInt();
            lcbPlcfBkfFactoid = stream.ReadInt();
            fcPlcpublicfcookie = stream.ReadInt();
            lcbPlcpublicfcookie = stream.ReadInt();
            fcPlcfBklFactoid = stream.ReadInt();
            lcbPlcfBklFactoid = stream.ReadInt();
            fcFactoidData = stream.ReadInt();
            lcbFactoidData = stream.ReadInt();
            fcDocUndo = stream.ReadInt();
            lcbDocUndo = stream.ReadInt();
            fcSttbfBkmkFcc = stream.ReadInt();
            lcbSttbfBkmkFcc = stream.ReadInt();
            fcPlcfBkfFcc = stream.ReadInt();
            lcbPlcfBkfFcc = stream.ReadInt();
            fcPlcfBklFcc = stream.ReadInt();
            lcbPlcfBklFcc = stream.ReadInt();
            fcSttbfbkmkBPRepairs = stream.ReadInt();
            lcbSttbfbkmkBPRepairs = stream.ReadInt();
            fcPlcfbkfBPRepairs = stream.ReadInt();
            lcbPlcfbkfBPRepairs = stream.ReadInt();
            fcPlcfbklBPRepairs = stream.ReadInt();
            lcbPlcfbklBPRepairs = stream.ReadInt();
            fcPmsNew = stream.ReadInt();
            lcbPmsNew = stream.ReadInt();
            fcODSO = stream.ReadInt();
            lcbODSO = stream.ReadInt();
            fcPlcfpmiOldXP = stream.ReadInt();
            lcbPlcfpmiOldXP = stream.ReadInt();
            fcPlcfpmiNewXP = stream.ReadInt();
            lcbPlcfpmiNewXP = stream.ReadInt();
            fcPlcfpmiMixedXP = stream.ReadInt();
            lcbPlcfpmiMixedXP = stream.ReadInt();
            fcUnused2 = stream.ReadInt();
            lcbUnused2 = stream.ReadInt();
            fcPlcffactoid = stream.ReadInt();
            lcbPlcffactoid = stream.ReadInt();
            fcPlcflvcOldXP = stream.ReadInt();
            lcbPlcflvcOldXP = stream.ReadInt();
            fcPlcflvcNewXP = stream.ReadInt();
            lcbPlcflvcNewXP = stream.ReadInt();
            fcPlcflvcMixedXP = stream.ReadInt();
            lcbPlcflvcMixedXP = stream.ReadInt();
        }
    }
}

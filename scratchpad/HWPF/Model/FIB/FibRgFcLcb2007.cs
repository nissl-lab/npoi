using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FibRgFcLcb2007
    {
        public int fcPlcfmthd;
        public int lcbPlcfmthd;
        public int fcSttbfBkmkMoveFrom;
        public int lcbSttbfBkmkMoveFrom;
        public int fcPlcfBkfMoveFrom;
        public int lcbPlcfBkfMoveFrom;
        public int fcPlcfBklMoveFrom;
        public int lcbPlcfBklMoveFrom;
        public int fcSttbfBkmkMoveTo;
        public int lcbSttbfBkmkMoveTo;
        public int fcPlcfBkfMoveTo;
        public int lcbPlcfBkfMoveTo;
        public int fcPlcfBklMoveTo;
        public int lcbPlcfBklMoveTo;
        public int fcUnused1;
        public int lcbUnused1;
        public int fcUnused2;
        public int lcbUnused2;
        public int fcUnused3;
        public int lcbUnused3;
        public int fcSttbfBkmkArto;
        public int lcbSttbfBkmkArto;
        public int fcPlcfBkfArto;
        public int lcbPlcfBkfArto;
        public int fcPlcfBklArto;
        public int lcbPlcfBklArto;
        public int fcArtoData;
        public int lcbArtoData;
        public int fcUnused4;
        public int lcbUnused4;
        public int fcUnused5;
        public int lcbUnused5;
        public int fcUnused6;
        public int lcbUnused6;
        public int fcOssTheme;
        public int lcbOssTheme;
        public int fcColorSchemeMapping;
        public int lcbColorSchemeMapping;

        public FibRgFcLcb2003 fibRgFcLcb2003;

        public FibRgFcLcb2007()
        {
        }
        public void Deserialize(HWPFStream stream)
        {
            fibRgFcLcb2003 = new FibRgFcLcb2003();
            fibRgFcLcb2003.Deserialize(stream);

            fcPlcfmthd = stream.ReadInt();
            lcbPlcfmthd = stream.ReadInt();
            fcSttbfBkmkMoveFrom = stream.ReadInt();
            lcbSttbfBkmkMoveFrom = stream.ReadInt();
            fcPlcfBkfMoveFrom = stream.ReadInt();
            lcbPlcfBkfMoveFrom = stream.ReadInt();
            fcPlcfBklMoveFrom = stream.ReadInt();
            lcbPlcfBklMoveFrom = stream.ReadInt();
            fcSttbfBkmkMoveTo = stream.ReadInt();
            lcbSttbfBkmkMoveTo = stream.ReadInt();
            fcPlcfBkfMoveTo = stream.ReadInt();
            lcbPlcfBkfMoveTo = stream.ReadInt();
            fcPlcfBklMoveTo = stream.ReadInt();
            lcbPlcfBklMoveTo = stream.ReadInt();
            fcUnused1 = stream.ReadInt();
            lcbUnused1 = stream.ReadInt();
            fcUnused2 = stream.ReadInt();
            lcbUnused2 = stream.ReadInt();
            fcUnused3 = stream.ReadInt();
            lcbUnused3 = stream.ReadInt();
            fcSttbfBkmkArto = stream.ReadInt();
            lcbSttbfBkmkArto = stream.ReadInt();
            fcPlcfBkfArto = stream.ReadInt();
            lcbPlcfBkfArto = stream.ReadInt();
            fcPlcfBklArto = stream.ReadInt();
            lcbPlcfBklArto = stream.ReadInt();
            fcArtoData = stream.ReadInt();
            lcbArtoData = stream.ReadInt();
            fcUnused4 = stream.ReadInt();
            lcbUnused4 = stream.ReadInt();
            fcUnused5 = stream.ReadInt();
            lcbUnused5 = stream.ReadInt();
            fcUnused6 = stream.ReadInt();
            lcbUnused6 = stream.ReadInt();
            fcOssTheme = stream.ReadInt();
            lcbOssTheme = stream.ReadInt();
            fcColorSchemeMapping = stream.ReadInt();
            lcbColorSchemeMapping = stream.ReadInt();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;

namespace NPOI.HWPF.Model
{
    public enum SprmType : short
    {
        sprmClbcCRJ = 0x00,
        sprmPIncLvl = 0x04,
        sprmPJc = 0x05,
        sprmPFKeep = 0x07,
        sprmPFKeepFollow = 0x08,
        sprmPFPageBreakBefore = 0x09,
        sprmPIlvl = 0x0C,
        sprmPFMirrorIndents = 0x0D,
        sprmPFNoLineNumb = 0x0E,
        sprmPTtwo = 0x0F,
        sprmPFInTable = 0x18,
        sprmPFTtp = 0x19,
        sprmPPc = 0x1D,
        sprmPWr = 0x25,
        sprmPFNoAutoHyph = 0x2C,
        sprmPFLocked = 0x32,
        sprmPFWidowControl = 0x33,
        sprmPFKinsoku = 0x35,
        sprmPFWordWrap = 0x36,
        sprmPFOverflowPunct = 0x37,
        sprmPFTopLinePunct = 0x38,
        sprmPFAutoSpaceDE = 0x39,
        sprmPFAutoSpaceDN = 0x3A,
        sprmCFRMarkDel = 0x41,
        sprmCFRMarkIns = 0x42,
        sprmCFFldVanish = 0x43,
        sprmCFData = 0x47,
        sprmCFOle2 = 0x4B,
        sprmCHighlight = 0x4D,
        sprmCFEmboss = 0x4E,
        sprmCSfxText = 0x4F,
        sprmCFWebHidden = 0x50,
        sprmCFSpecVanish = 0x51,
        sprmCPlain = 0x53,
        sprmCFBold = 0x55,
        sprmCFItalic = 0x56,
        sprmCFStrike = 0x57,
        sprmCFOutline = 0x58,
        sprmCFShadow = 0x59,
        sprmCFSmallCaps = 0x5A,
        sprmCFCaps = 0x5B,
        sprmCFVanish = 0x5C,
        sprmCKul = 0x5E,
    }

    public class Prm0
    {
        private byte field_1_fComplex = 0;
        private static BitField field_2_isprm = BitFieldFactory.GetInstance(0x7F);

    }
}

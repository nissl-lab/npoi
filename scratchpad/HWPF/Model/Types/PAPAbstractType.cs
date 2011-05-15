
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */



namespace NPOI.HWPF.Model.Types
{
    using System;
    using NPOI.Util;

    using NPOI.HWPF.UserModel;
    using System.Text;


    /**
     * Paragraph Properties.
     * NOTE: This source Is automatically generated please do not modify this file.  Either subclass or
     *       remove the record in src/records/definitions.

     * @author S. Ryan Ackley
     */
    public abstract class PAPAbstractType: BaseObject
    {

        protected int field_1_istd;
        protected byte field_2_jc;
        protected byte field_3_fKeep;
        protected byte field_4_fKeepFollow;
        protected byte field_5_fPageBreakBefore;
        protected byte field_6_fBrLnAbove;
        protected byte field_7_fBrLnBelow;
        protected byte field_8_pcVert;
        protected byte field_9_pcHorz;
        protected byte field_10_brcp;
        protected byte field_11_brcl;
        protected byte field_12_ilvl;
        protected byte field_13_fNoLnn;
        protected int field_14_ilfo;
        protected byte field_15_fSideBySide;
        protected byte field_16_fNoAutoHyph;
        protected byte field_17_fWidowControl;
        protected int field_18_dxaRight;
        protected int field_19_dxaLeft;
        protected int field_20_dxaLeft1;
        protected LineSpacingDescriptor field_21_lspd;
        protected int field_22_dyaBefore;
        protected int field_23_dyaAfter;
        protected byte[] field_24_phe;
        protected byte field_25_fCrLf;
        protected byte field_26_fUsePgsuSettings;
        protected byte field_27_fAdjustRight;
        protected byte field_28_fKinsoku;
        protected byte field_29_fWordWrap;
        protected byte field_30_fOverflowPunct;
        protected byte field_31_fTopLinePunct;
        protected byte field_32_fAutoSpaceDE;
        protected byte field_33_fAutoSpaceDN;
        protected int field_34_wAlignFont;
        protected short field_35_fontAlign;
        private static BitField fVertical = BitFieldFactory.GetInstance(0x0001);
        private static BitField fBackward = BitFieldFactory.GetInstance(0x0002);
        private static BitField fRotateFont = BitFieldFactory.GetInstance(0x0004);
        protected byte field_36_fBackward;
        protected byte field_37_fRotateFont;
        protected byte field_38_fInTable;
        protected byte field_39_fTtp;
        protected byte field_40_wr;
        protected byte field_41_fLocked;
        protected byte[] field_42_ptap;
        protected int field_43_dxaAbs;
        protected int field_44_dyaAbs;
        protected int field_45_dxaWidth;
        protected BorderCode field_46_brcTop;
        protected BorderCode field_47_brcLeft;
        protected BorderCode field_48_brcBottom;
        protected BorderCode field_49_brcRight;
        protected BorderCode field_50_brcBetween;
        protected BorderCode field_51_brcBar;
        protected int field_52_dxaFromText;
        protected int field_53_dyaFromText;
        protected int field_54_dyaHeight;
        protected byte field_55_fMinHeight;
        protected ShadingDescriptor field_56_shd;
        protected DropCapSpecifier field_57_dcs;
        protected byte field_58_lvl = 9;
        protected byte field_59_fNumRMIns;
        protected byte[] field_60_anld;
        protected int field_61_fPropRMark;
        protected int field_62_ibstPropRMark;
        protected DateAndTime field_63_dttmPropRMark;
        protected byte[] field_64_numrm;
        protected int field_65_itbdMac;
        protected int[] field_66_rgdxaTab;
        protected byte[] field_67_rgtbd;
        protected byte field_68_tableLevel;
        protected byte field_69_fTtpEmbedded;
        protected byte field_70_embeddedCellMark;


        public PAPAbstractType()
        {

        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public int GetSize()
        {
            return 4 + +2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 2 + 1 + 1 + 1 + 4 + 4 + 4 + 4 + 4 + 4 + 12 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 2 + 1 + 2 + 2 + 1 + 1 + 84 + 1 + 2 + 4 + 128 + 2 + 128 + 128 + 1 + 1 + 1;
        }



        /**
         * Get the Istd field for the PAP record.
         */
        public int GetIstd()
        {
            return field_1_istd;
        }

        /**
         * Set the Istd field for the PAP record.
         */
        public void SetIstd(int field_1_istd)
        {
            this.field_1_istd = field_1_istd;
        }

        /**
         * Get the jc field for the PAP record.
         */
        public byte GetJc()
        {
            return field_2_jc;
        }

        /**
         * Set the jc field for the PAP record.
         */
        public void SetJc(byte field_2_jc)
        {
            this.field_2_jc = field_2_jc;
        }

        /**
         * Get the fKeep field for the PAP record.
         */
        public byte GetFKeep()
        {
            return field_3_fKeep;
        }

        /**
         * Set the fKeep field for the PAP record.
         */
        public void SetFKeep(byte field_3_fKeep)
        {
            this.field_3_fKeep = field_3_fKeep;
        }

        /**
         * Get the fKeepFollow field for the PAP record.
         */
        public byte GetFKeepFollow()
        {
            return field_4_fKeepFollow;
        }

        /**
         * Set the fKeepFollow field for the PAP record.
         */
        public void SetFKeepFollow(byte field_4_fKeepFollow)
        {
            this.field_4_fKeepFollow = field_4_fKeepFollow;
        }

        /**
         * Get the fPageBreakBefore field for the PAP record.
         */
        public byte GetFPageBreakBefore()
        {
            return field_5_fPageBreakBefore;
        }

        /**
         * Set the fPageBreakBefore field for the PAP record.
         */
        public void SetFPageBreakBefore(byte field_5_fPageBreakBefore)
        {
            this.field_5_fPageBreakBefore = field_5_fPageBreakBefore;
        }

        /**
         * Get the fBrLnAbove field for the PAP record.
         */
        public byte GetFBrLnAbove()
        {
            return field_6_fBrLnAbove;
        }

        /**
         * Set the fBrLnAbove field for the PAP record.
         */
        public void SetFBrLnAbove(byte field_6_fBrLnAbove)
        {
            this.field_6_fBrLnAbove = field_6_fBrLnAbove;
        }

        /**
         * Get the fBrLnBelow field for the PAP record.
         */
        public byte GetFBrLnBelow()
        {
            return field_7_fBrLnBelow;
        }

        /**
         * Set the fBrLnBelow field for the PAP record.
         */
        public void SetFBrLnBelow(byte field_7_fBrLnBelow)
        {
            this.field_7_fBrLnBelow = field_7_fBrLnBelow;
        }

        /**
         * Get the pcVert field for the PAP record.
         */
        public byte GetPcVert()
        {
            return field_8_pcVert;
        }

        /**
         * Set the pcVert field for the PAP record.
         */
        public void SetPcVert(byte field_8_pcVert)
        {
            this.field_8_pcVert = field_8_pcVert;
        }

        /**
         * Get the pcHorz field for the PAP record.
         */
        public byte GetPcHorz()
        {
            return field_9_pcHorz;
        }

        /**
         * Set the pcHorz field for the PAP record.
         */
        public void SetPcHorz(byte field_9_pcHorz)
        {
            this.field_9_pcHorz = field_9_pcHorz;
        }

        /**
         * Get the brcp field for the PAP record.
         */
        public byte GetBrcp()
        {
            return field_10_brcp;
        }

        /**
         * Set the brcp field for the PAP record.
         */
        public void SetBrcp(byte field_10_brcp)
        {
            this.field_10_brcp = field_10_brcp;
        }

        /**
         * Get the brcl field for the PAP record.
         */
        public byte GetBrcl()
        {
            return field_11_brcl;
        }

        /**
         * Set the brcl field for the PAP record.
         */
        public void SetBrcl(byte field_11_brcl)
        {
            this.field_11_brcl = field_11_brcl;
        }

        /**
         * Get the ilvl field for the PAP record.
         */
        public byte GetIlvl()
        {
            return field_12_ilvl;
        }

        /**
         * Set the ilvl field for the PAP record.
         */
        public void SetIlvl(byte field_12_ilvl)
        {
            this.field_12_ilvl = field_12_ilvl;
        }

        /**
         * Get the fNoLnn field for the PAP record.
         */
        public byte GetFNoLnn()
        {
            return field_13_fNoLnn;
        }

        /**
         * Set the fNoLnn field for the PAP record.
         */
        public void SetFNoLnn(byte field_13_fNoLnn)
        {
            this.field_13_fNoLnn = field_13_fNoLnn;
        }

        /**
         * Get the ilfo field for the PAP record.
         */
        public int GetIlfo()
        {
            return field_14_ilfo;
        }

        /**
         * Set the ilfo field for the PAP record.
         */
        public void SetIlfo(int field_14_ilfo)
        {
            this.field_14_ilfo = field_14_ilfo;
        }

        /**
         * Get the fSideBySide field for the PAP record.
         */
        public byte GetFSideBySide()
        {
            return field_15_fSideBySide;
        }

        /**
         * Set the fSideBySide field for the PAP record.
         */
        public void SetFSideBySide(byte field_15_fSideBySide)
        {
            this.field_15_fSideBySide = field_15_fSideBySide;
        }

        /**
         * Get the fNoAutoHyph field for the PAP record.
         */
        public byte GetFNoAutoHyph()
        {
            return field_16_fNoAutoHyph;
        }

        /**
         * Set the fNoAutoHyph field for the PAP record.
         */
        public void SetFNoAutoHyph(byte field_16_fNoAutoHyph)
        {
            this.field_16_fNoAutoHyph = field_16_fNoAutoHyph;
        }

        /**
         * Get the fWidowControl field for the PAP record.
         */
        public byte GetFWidowControl()
        {
            return field_17_fWidowControl;
        }

        /**
         * Set the fWidowControl field for the PAP record.
         */
        public void SetFWidowControl(byte field_17_fWidowControl)
        {
            this.field_17_fWidowControl = field_17_fWidowControl;
        }

        /**
         * Get the dxaRight field for the PAP record.
         */
        public int GetDxaRight()
        {
            return field_18_dxaRight;
        }

        /**
         * Set the dxaRight field for the PAP record.
         */
        public void SetDxaRight(int field_18_dxaRight)
        {
            this.field_18_dxaRight = field_18_dxaRight;
        }

        /**
         * Get the dxaLeft field for the PAP record.
         */
        public int GetDxaLeft()
        {
            return field_19_dxaLeft;
        }

        /**
         * Set the dxaLeft field for the PAP record.
         */
        public void SetDxaLeft(int field_19_dxaLeft)
        {
            this.field_19_dxaLeft = field_19_dxaLeft;
        }

        /**
         * Get the dxaLeft1 field for the PAP record.
         */
        public int GetDxaLeft1()
        {
            return field_20_dxaLeft1;
        }

        /**
         * Set the dxaLeft1 field for the PAP record.
         */
        public void SetDxaLeft1(int field_20_dxaLeft1)
        {
            this.field_20_dxaLeft1 = field_20_dxaLeft1;
        }

        /**
         * Get the lspd field for the PAP record.
         */
        public LineSpacingDescriptor GetLspd()
        {
            return field_21_lspd;
        }

        /**
         * Set the lspd field for the PAP record.
         */
        public void SetLspd(LineSpacingDescriptor field_21_lspd)
        {
            this.field_21_lspd = field_21_lspd;
        }

        /**
         * Get the dyaBefore field for the PAP record.
         */
        public int GetDyaBefore()
        {
            return field_22_dyaBefore;
        }

        /**
         * Set the dyaBefore field for the PAP record.
         */
        public void SetDyaBefore(int field_22_dyaBefore)
        {
            this.field_22_dyaBefore = field_22_dyaBefore;
        }

        /**
         * Get the dyaAfter field for the PAP record.
         */
        public int GetDyaAfter()
        {
            return field_23_dyaAfter;
        }

        /**
         * Set the dyaAfter field for the PAP record.
         */
        public void SetDyaAfter(int field_23_dyaAfter)
        {
            this.field_23_dyaAfter = field_23_dyaAfter;
        }

        /**
         * Get the phe field for the PAP record.
         */
        public byte[] GetPhe()
        {
            return field_24_phe;
        }

        /**
         * Set the phe field for the PAP record.
         */
        public void SetPhe(byte[] field_24_phe)
        {
            this.field_24_phe = field_24_phe;
        }

        /**
         * Get the fCrLf field for the PAP record.
         */
        public byte GetFCrLf()
        {
            return field_25_fCrLf;
        }

        /**
         * Set the fCrLf field for the PAP record.
         */
        public void SetFCrLf(byte field_25_fCrLf)
        {
            this.field_25_fCrLf = field_25_fCrLf;
        }

        /**
         * Get the fUsePgsuSettings field for the PAP record.
         */
        public byte GetFUsePgsuSettings()
        {
            return field_26_fUsePgsuSettings;
        }

        /**
         * Set the fUsePgsuSettings field for the PAP record.
         */
        public void SetFUsePgsuSettings(byte field_26_fUsePgsuSettings)
        {
            this.field_26_fUsePgsuSettings = field_26_fUsePgsuSettings;
        }

        /**
         * Get the fAdjustRight field for the PAP record.
         */
        public byte GetFAdjustRight()
        {
            return field_27_fAdjustRight;
        }

        /**
         * Set the fAdjustRight field for the PAP record.
         */
        public void SetFAdjustRight(byte field_27_fAdjustRight)
        {
            this.field_27_fAdjustRight = field_27_fAdjustRight;
        }

        /**
         * Get the fKinsoku field for the PAP record.
         */
        public byte GetFKinsoku()
        {
            return field_28_fKinsoku;
        }

        /**
         * Set the fKinsoku field for the PAP record.
         */
        public void SetFKinsoku(byte field_28_fKinsoku)
        {
            this.field_28_fKinsoku = field_28_fKinsoku;
        }

        /**
         * Get the fWordWrap field for the PAP record.
         */
        public byte GetFWordWrap()
        {
            return field_29_fWordWrap;
        }

        /**
         * Set the fWordWrap field for the PAP record.
         */
        public void SetFWordWrap(byte field_29_fWordWrap)
        {
            this.field_29_fWordWrap = field_29_fWordWrap;
        }

        /**
         * Get the fOverflowPunct field for the PAP record.
         */
        public byte GetFOverflowPunct()
        {
            return field_30_fOverflowPunct;
        }

        /**
         * Set the fOverflowPunct field for the PAP record.
         */
        public void SetFOverflowPunct(byte field_30_fOverflowPunct)
        {
            this.field_30_fOverflowPunct = field_30_fOverflowPunct;
        }

        /**
         * Get the fTopLinePunct field for the PAP record.
         */
        public byte GetFTopLinePunct()
        {
            return field_31_fTopLinePunct;
        }

        /**
         * Set the fTopLinePunct field for the PAP record.
         */
        public void SetFTopLinePunct(byte field_31_fTopLinePunct)
        {
            this.field_31_fTopLinePunct = field_31_fTopLinePunct;
        }

        /**
         * Get the fAutoSpaceDE field for the PAP record.
         */
        public byte GetFAutoSpaceDE()
        {
            return field_32_fAutoSpaceDE;
        }

        /**
         * Set the fAutoSpaceDE field for the PAP record.
         */
        public void SetFAutoSpaceDE(byte field_32_fAutoSpaceDE)
        {
            this.field_32_fAutoSpaceDE = field_32_fAutoSpaceDE;
        }

        /**
         * Get the fAutoSpaceDN field for the PAP record.
         */
        public byte GetFAutoSpaceDN()
        {
            return field_33_fAutoSpaceDN;
        }

        /**
         * Set the fAutoSpaceDN field for the PAP record.
         */
        public void SetFAutoSpaceDN(byte field_33_fAutoSpaceDN)
        {
            this.field_33_fAutoSpaceDN = field_33_fAutoSpaceDN;
        }

        /**
         * Get the wAlignFont field for the PAP record.
         */
        public int GetWAlignFont()
        {
            return field_34_wAlignFont;
        }

        /**
         * Set the wAlignFont field for the PAP record.
         */
        public void SetWAlignFont(int field_34_wAlignFont)
        {
            this.field_34_wAlignFont = field_34_wAlignFont;
        }

        /**
         * Get the fontAlign field for the PAP record.
         */
        public short GetFontAlign()
        {
            return field_35_fontAlign;
        }

        /**
         * Set the fontAlign field for the PAP record.
         */
        public void SetFontAlign(short field_35_fontAlign)
        {
            this.field_35_fontAlign = field_35_fontAlign;
        }

        /**
         * Get the fBackward field for the PAP record.
         */
        public byte GetFBackward()
        {
            return field_36_fBackward;
        }

        /**
         * Set the fBackward field for the PAP record.
         */
        public void SetFBackward(byte field_36_fBackward)
        {
            this.field_36_fBackward = field_36_fBackward;
        }

        /**
         * Get the fRotateFont field for the PAP record.
         */
        public byte GetFRotateFont()
        {
            return field_37_fRotateFont;
        }

        /**
         * Set the fRotateFont field for the PAP record.
         */
        public void SetFRotateFont(byte field_37_fRotateFont)
        {
            this.field_37_fRotateFont = field_37_fRotateFont;
        }

        /**
         * Get the fInTable field for the PAP record.
         */
        public byte GetFInTable()
        {
            return field_38_fInTable;
        }

        /**
         * Set the fInTable field for the PAP record.
         */
        public void SetFInTable(byte field_38_fInTable)
        {
            this.field_38_fInTable = field_38_fInTable;
        }

        /**
         * Get the fTtp field for the PAP record.
         */
        public byte GetFTtp()
        {
            return field_39_fTtp;
        }

        /**
         * Set the fTtp field for the PAP record.
         */
        public void SetFTtp(byte field_39_fTtp)
        {
            this.field_39_fTtp = field_39_fTtp;
        }

        /**
         * Get the wr field for the PAP record.
         */
        public byte GetWr()
        {
            return field_40_wr;
        }

        /**
         * Set the wr field for the PAP record.
         */
        public void SetWr(byte field_40_wr)
        {
            this.field_40_wr = field_40_wr;
        }

        /**
         * Get the fLocked field for the PAP record.
         */
        public byte GetFLocked()
        {
            return field_41_fLocked;
        }

        /**
         * Set the fLocked field for the PAP record.
         */
        public void SetFLocked(byte field_41_fLocked)
        {
            this.field_41_fLocked = field_41_fLocked;
        }

        /**
         * Get the ptap field for the PAP record.
         */
        public byte[] GetPtap()
        {
            return field_42_ptap;
        }

        /**
         * Set the ptap field for the PAP record.
         */
        public void SetPtap(byte[] field_42_ptap)
        {
            this.field_42_ptap = field_42_ptap;
        }

        /**
         * Get the dxaAbs field for the PAP record.
         */
        public int GetDxaAbs()
        {
            return field_43_dxaAbs;
        }

        /**
         * Set the dxaAbs field for the PAP record.
         */
        public void SetDxaAbs(int field_43_dxaAbs)
        {
            this.field_43_dxaAbs = field_43_dxaAbs;
        }

        /**
         * Get the dyaAbs field for the PAP record.
         */
        public int GetDyaAbs()
        {
            return field_44_dyaAbs;
        }

        /**
         * Set the dyaAbs field for the PAP record.
         */
        public void SetDyaAbs(int field_44_dyaAbs)
        {
            this.field_44_dyaAbs = field_44_dyaAbs;
        }

        /**
         * Get the dxaWidth field for the PAP record.
         */
        public int GetDxaWidth()
        {
            return field_45_dxaWidth;
        }

        /**
         * Set the dxaWidth field for the PAP record.
         */
        public void SetDxaWidth(int field_45_dxaWidth)
        {
            this.field_45_dxaWidth = field_45_dxaWidth;
        }

        /**
         * Get the brcTop field for the PAP record.
         */
        public BorderCode GetBrcTop()
        {
            return field_46_brcTop;
        }

        /**
         * Set the brcTop field for the PAP record.
         */
        public void SetBrcTop(BorderCode field_46_brcTop)
        {
            this.field_46_brcTop = field_46_brcTop;
        }

        /**
         * Get the brcLeft field for the PAP record.
         */
        public BorderCode GetBrcLeft()
        {
            return field_47_brcLeft;
        }

        /**
         * Set the brcLeft field for the PAP record.
         */
        public void SetBrcLeft(BorderCode field_47_brcLeft)
        {
            this.field_47_brcLeft = field_47_brcLeft;
        }

        /**
         * Get the brcBottom field for the PAP record.
         */
        public BorderCode GetBrcBottom()
        {
            return field_48_brcBottom;
        }

        /**
         * Set the brcBottom field for the PAP record.
         */
        public void SetBrcBottom(BorderCode field_48_brcBottom)
        {
            this.field_48_brcBottom = field_48_brcBottom;
        }

        /**
         * Get the brcRight field for the PAP record.
         */
        public BorderCode GetBrcRight()
        {
            return field_49_brcRight;
        }

        /**
         * Set the brcRight field for the PAP record.
         */
        public void SetBrcRight(BorderCode field_49_brcRight)
        {
            this.field_49_brcRight = field_49_brcRight;
        }

        /**
         * Get the brcBetween field for the PAP record.
         */
        public BorderCode GetBrcBetween()
        {
            return field_50_brcBetween;
        }

        /**
         * Set the brcBetween field for the PAP record.
         */
        public void SetBrcBetween(BorderCode field_50_brcBetween)
        {
            this.field_50_brcBetween = field_50_brcBetween;
        }

        /**
         * Get the brcBar field for the PAP record.
         */
        public BorderCode GetBrcBar()
        {
            return field_51_brcBar;
        }

        /**
         * Set the brcBar field for the PAP record.
         */
        public void SetBrcBar(BorderCode field_51_brcBar)
        {
            this.field_51_brcBar = field_51_brcBar;
        }

        /**
         * Get the dxaFromText field for the PAP record.
         */
        public int GetDxaFromText()
        {
            return field_52_dxaFromText;
        }

        /**
         * Set the dxaFromText field for the PAP record.
         */
        public void SetDxaFromText(int field_52_dxaFromText)
        {
            this.field_52_dxaFromText = field_52_dxaFromText;
        }

        /**
         * Get the dyaFromText field for the PAP record.
         */
        public int GetDyaFromText()
        {
            return field_53_dyaFromText;
        }

        /**
         * Set the dyaFromText field for the PAP record.
         */
        public void SetDyaFromText(int field_53_dyaFromText)
        {
            this.field_53_dyaFromText = field_53_dyaFromText;
        }

        /**
         * Get the dyaHeight field for the PAP record.
         */
        public int GetDyaHeight()
        {
            return field_54_dyaHeight;
        }

        /**
         * Set the dyaHeight field for the PAP record.
         */
        public void SetDyaHeight(int field_54_dyaHeight)
        {
            this.field_54_dyaHeight = field_54_dyaHeight;
        }

        /**
         * Get the fMinHeight field for the PAP record.
         */
        public byte GetFMinHeight()
        {
            return field_55_fMinHeight;
        }

        /**
         * Set the fMinHeight field for the PAP record.
         */
        public void SetFMinHeight(byte field_55_fMinHeight)
        {
            this.field_55_fMinHeight = field_55_fMinHeight;
        }

        /**
         * Get the shd field for the PAP record.
         */
        public ShadingDescriptor GetShd()
        {
            return field_56_shd;
        }

        /**
         * Set the shd field for the PAP record.
         */
        public void SetShd(ShadingDescriptor field_56_shd)
        {
            this.field_56_shd = field_56_shd;
        }

        /**
         * Get the dcs field for the PAP record.
         */
        public DropCapSpecifier GetDcs()
        {
            return field_57_dcs;
        }

        /**
         * Set the dcs field for the PAP record.
         */
        public void SetDcs(DropCapSpecifier field_57_dcs)
        {
            this.field_57_dcs = field_57_dcs;
        }

        /**
         * Get the lvl field for the PAP record.
         */
        public byte GetLvl()
        {
            return field_58_lvl;
        }

        /**
         * Set the lvl field for the PAP record.
         */
        public void SetLvl(byte field_58_lvl)
        {
            this.field_58_lvl = field_58_lvl;
        }

        /**
         * Get the fNumRMIns field for the PAP record.
         */
        public byte GetFNumRMIns()
        {
            return field_59_fNumRMIns;
        }

        /**
         * Set the fNumRMIns field for the PAP record.
         */
        public void SetFNumRMIns(byte field_59_fNumRMIns)
        {
            this.field_59_fNumRMIns = field_59_fNumRMIns;
        }

        /**
         * Get the anld field for the PAP record.
         */
        public byte[] GetAnld()
        {
            return field_60_anld;
        }

        /**
         * Set the anld field for the PAP record.
         */
        public void SetAnld(byte[] field_60_anld)
        {
            this.field_60_anld = field_60_anld;
        }

        /**
         * Get the fPropRMark field for the PAP record.
         */
        public int GetFPropRMark()
        {
            return field_61_fPropRMark;
        }

        /**
         * Set the fPropRMark field for the PAP record.
         */
        public void SetFPropRMark(int field_61_fPropRMark)
        {
            this.field_61_fPropRMark = field_61_fPropRMark;
        }

        /**
         * Get the ibstPropRMark field for the PAP record.
         */
        public int GetIbstPropRMark()
        {
            return field_62_ibstPropRMark;
        }

        /**
         * Set the ibstPropRMark field for the PAP record.
         */
        public void SetIbstPropRMark(int field_62_ibstPropRMark)
        {
            this.field_62_ibstPropRMark = field_62_ibstPropRMark;
        }

        /**
         * Get the dttmPropRMark field for the PAP record.
         */
        public DateAndTime GetDttmPropRMark()
        {
            return field_63_dttmPropRMark;
        }

        /**
         * Set the dttmPropRMark field for the PAP record.
         */
        public void SetDttmPropRMark(DateAndTime field_63_dttmPropRMark)
        {
            this.field_63_dttmPropRMark = field_63_dttmPropRMark;
        }

        /**
         * Get the numrm field for the PAP record.
         */
        public byte[] GetNumrm()
        {
            return field_64_numrm;
        }

        /**
         * Set the numrm field for the PAP record.
         */
        public void SetNumrm(byte[] field_64_numrm)
        {
            this.field_64_numrm = field_64_numrm;
        }

        /**
         * Get the itbdMac field for the PAP record.
         */
        public int GetItbdMac()
        {
            return field_65_itbdMac;
        }

        /**
         * Set the itbdMac field for the PAP record.
         */
        public void SetItbdMac(int field_65_itbdMac)
        {
            this.field_65_itbdMac = field_65_itbdMac;
        }

        /**
         * Get the rgdxaTab field for the PAP record.
         */
        public int[] GetRgdxaTab()
        {
            return field_66_rgdxaTab;
        }

        /**
         * Set the rgdxaTab field for the PAP record.
         */
        public void SetRgdxaTab(int[] field_66_rgdxaTab)
        {
            this.field_66_rgdxaTab = field_66_rgdxaTab;
        }

        /**
         * Get the rgtbd field for the PAP record.
         */
        public byte[] GetRgtbd()
        {
            return field_67_rgtbd;
        }

        /**
         * Set the rgtbd field for the PAP record.
         */
        public void SetRgtbd(byte[] field_67_rgtbd)
        {
            this.field_67_rgtbd = field_67_rgtbd;
        }

        /**
         * Get the tableLevel field for the PAP record.
         */
        public byte GetTableLevel()
        {
            return field_68_tableLevel;
        }

        /**
         * Set the tableLevel field for the PAP record.
         */
        public void SetTableLevel(byte field_68_tableLevel)
        {
            this.field_68_tableLevel = field_68_tableLevel;
        }

        /**
         * Get the fTtpEmbedded field for the PAP record.
         */
        public byte GetFTtpEmbedded()
        {
            return field_69_fTtpEmbedded;
        }

        /**
         * Set the fTtpEmbedded field for the PAP record.
         */
        public void SetFTtpEmbedded(byte field_69_fTtpEmbedded)
        {
            this.field_69_fTtpEmbedded = field_69_fTtpEmbedded;
        }

        /**
         * Get the embeddedCellMark field for the PAP record.
         */
        public byte GetEmbeddedCellMark()
        {
            return field_70_embeddedCellMark;
        }

        /**
         * Set the embeddedCellMark field for the PAP record.
         */
        public void SetEmbeddedCellMark(byte field_70_embeddedCellMark)
        {
            this.field_70_embeddedCellMark = field_70_embeddedCellMark;
        }

        /**
         * Sets the fVertical field value.
         * 
         */
        public void SetFVertical(bool value)
        {
            field_35_fontAlign = (short)fVertical.SetBoolean(field_35_fontAlign, value);


        }

        /**
         * 
         * @return  the fVertical field value.
         */
        public bool IsFVertical()
        {
            return fVertical.IsSet(field_35_fontAlign);

        }

        /**
         * Sets the fBackward field value.
         * 
         */
        public void SetFBackward(bool value)
        {
            field_35_fontAlign = (short)fBackward.SetBoolean(field_35_fontAlign, value);


        }

        /**
         * 
         * @return  the fBackward field value.
         */
        public bool IsFBackward()
        {
            return fBackward.IsSet(field_35_fontAlign);

        }

        /**
         * Sets the fRotateFont field value.
         * 
         */
        public void SetFRotateFont(bool value)
        {
            field_35_fontAlign = (short)fRotateFont.SetBoolean(field_35_fontAlign, value);


        }

        /**
         * 
         * @return  the fRotateFont field value.
         */
        public bool IsFRotateFont()
        {
            return fRotateFont.IsSet(field_35_fontAlign);

        }


    }
}
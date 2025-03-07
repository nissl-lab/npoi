/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.Model.Types
{

    using NPOI.HWPF.UserModel;
    using NPOI.Util;
    using System;
    using System.Text;

    /**
     * Paragraph Properties.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       remove the record in src/records/defInitions.

     * @author S. Ryan Ackley
     */
    public abstract class PAPAbstractType:BaseObject
    {

        protected int field_1_istd;
        protected bool field_2_fSideBySide;
        protected bool field_3_fKeep;
        protected bool field_4_fKeepFollow;
        protected bool field_5_fPageBreakBefore;
        protected byte field_6_brcl;
        /**/
        public static byte BRCL_SINGLE = 0;
        /**/
        public static byte BRCL_THICK = 1;
        /**/
        public static byte BRCL_DOUBLE = 2;
        /**/
        public static byte BRCL_SHADOW = 3;
        protected byte field_7_brcp;
        /**/
        public static byte BRCP_NONE = 0;
        /**/
        public static byte BRCP_BORDER_ABOVE = 1;
        /**/
        public static byte BRCP_BORDER_BELOW = 2;
        /**/
        public static byte BRCP_BOX_AROUND = 15;
        /**/
        public static byte BRCP_BAR_TO_LEFT_OF_PARAGRAPH = 16;
        protected byte field_8_ilvl;
        protected int field_9_ilfo;
        protected bool field_10_fNoLnn;
        protected LineSpacingDescriptor field_11_lspd;
        protected int field_12_dyaBefore;
        protected int field_13_dyaAfter;
        protected bool field_14_fInTable;
        protected bool field_15_finTableW97;
        protected bool field_16_fTtp;
        protected int field_17_dxaAbs;
        protected int field_18_dyaAbs;
        protected int field_19_dxaWidth;
        protected bool field_20_fBrLnAbove;
        protected bool field_21_fBrLnBelow;
        protected byte field_22_pcVert;
        protected byte field_23_pcHorz;
        protected byte field_24_wr;
        protected bool field_25_fNoAutoHyph;
        protected int field_26_dyaHeight;
        protected bool field_27_fMinHeight;
        /**/
        public static bool FMINHEIGHT_EXACT = false;
        /**/
        public static bool FMINHEIGHT_AT_LEAST = true;
        protected DropCapSpecifier field_28_dcs;
        protected int field_29_dyaFromText;
        protected int field_30_dxaFromText;
        protected bool field_31_fLocked;
        protected bool field_32_fWidowControl;
        protected bool field_33_fKinsoku;
        protected bool field_34_fWordWrap;
        protected bool field_35_fOverflowPunct;
        protected bool field_36_fTopLinePunct;
        protected bool field_37_fAutoSpaceDE;
        protected bool field_38_fAutoSpaceDN;
        protected int field_39_wAlignFont;
        /**/
        public static byte WALIGNFONT_HANGING = 0;
        /**/
        public static byte WALIGNFONT_CENTERED = 1;
        /**/
        public static byte WALIGNFONT_ROMAN = 2;
        /**/
        public static byte WALIGNFONT_VARIABLE = 3;
        /**/
        public static byte WALIGNFONT_AUTO = 4;
        protected short field_40_fontAlign;
        private static BitField fVertical = new BitField(0x0001);
        private static BitField fBackward = new BitField(0x0002);
        private static BitField fRotateFont = new BitField(0x0004);
        protected byte field_41_lvl;
        protected bool field_42_fBiDi;
        protected bool field_43_fNumRMIns;
        protected bool field_44_fCrLf;
        protected bool field_45_fUsePgsuSettings;
        protected bool field_46_fAdjustRight;
        protected int field_47_itap;
        protected bool field_48_fInnerTableCell;
        protected bool field_49_fOpenTch;
        protected bool field_50_fTtpEmbedded;
        protected short field_51_dxcRight;
        protected short field_52_dxcLeft;
        protected short field_53_dxcLeft1;
        protected bool field_54_fDyaBeforeAuto;
        protected bool field_55_fDyaAfterAuto;
        protected int field_56_dxaRight;
        protected int field_57_dxaLeft;
        protected int field_58_dxaLeft1;
        protected byte field_59_jc;
        protected bool field_60_fNoAllowOverlap;
        protected BorderCode field_61_brcTop;
        protected BorderCode field_62_brcLeft;
        protected BorderCode field_63_brcBottom;
        protected BorderCode field_64_brcRight;
        protected BorderCode field_65_brcBetween;
        protected BorderCode field_66_brcBar;
        protected ShadingDescriptor field_67_shd;
        protected byte[] field_68_anld;
        protected byte[] field_69_phe;
        protected bool field_70_fPropRMark;
        protected int field_71_ibstPropRMark;
        protected DateAndTime field_72_dttmPropRMark;
        protected int field_73_itbdMac;
        protected int[] field_74_rgdxaTab;
        protected byte[] field_75_rgtbd;
        protected byte[] field_76_numrm;
        protected byte[] field_77_ptap;

        protected PAPAbstractType()
        {
            this.field_11_lspd = new LineSpacingDescriptor();
            this.field_11_lspd = new LineSpacingDescriptor();
            this.field_28_dcs = new DropCapSpecifier();
            this.field_32_fWidowControl = true;
            this.field_41_lvl = 9;
            this.field_61_brcTop = new BorderCode();
            this.field_62_brcLeft = new BorderCode();
            this.field_63_brcBottom = new BorderCode();
            this.field_64_brcRight = new BorderCode();
            this.field_65_brcBetween = new BorderCode();
            this.field_66_brcBar = new BorderCode();
            this.field_67_shd = new ShadingDescriptor();
            this.field_68_anld = Array.Empty<byte>();
            this.field_69_phe = Array.Empty<byte>();
            this.field_72_dttmPropRMark = new DateAndTime();
            this.field_74_rgdxaTab = Array.Empty<int>();
            this.field_75_rgtbd = Array.Empty<byte>();
            this.field_76_numrm = Array.Empty<byte>();
            this.field_77_ptap = Array.Empty<byte>();
        }


        public override String ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("[PAP]\n");
            builder.Append("    .Istd                 = ");
            builder.Append(" (").Append(GetIstd()).Append(" )\n");
            builder.Append("    .fSideBySide          = ");
            builder.Append(" (").Append(GetFSideBySide()).Append(" )\n");
            builder.Append("    .fKeep                = ");
            builder.Append(" (").Append(GetFKeep()).Append(" )\n");
            builder.Append("    .fKeepFollow          = ");
            builder.Append(" (").Append(GetFKeepFollow()).Append(" )\n");
            builder.Append("    .fPageBreakBefore     = ");
            builder.Append(" (").Append(GetFPageBreakBefore()).Append(" )\n");
            builder.Append("    .brcl                 = ");
            builder.Append(" (").Append(GetBrcl()).Append(" )\n");
            builder.Append("    .brcp                 = ");
            builder.Append(" (").Append(GetBrcp()).Append(" )\n");
            builder.Append("    .ilvl                 = ");
            builder.Append(" (").Append(GetIlvl()).Append(" )\n");
            builder.Append("    .ilfo                 = ");
            builder.Append(" (").Append(GetIlfo()).Append(" )\n");
            builder.Append("    .fNoLnn               = ");
            builder.Append(" (").Append(GetFNoLnn()).Append(" )\n");
            builder.Append("    .lspd                 = ");
            builder.Append(" (").Append(GetLspd()).Append(" )\n");
            builder.Append("    .dyaBefore            = ");
            builder.Append(" (").Append(GetDyaBefore()).Append(" )\n");
            builder.Append("    .dyaAfter             = ");
            builder.Append(" (").Append(GetDyaAfter()).Append(" )\n");
            builder.Append("    .fInTable             = ");
            builder.Append(" (").Append(GetFInTable()).Append(" )\n");
            builder.Append("    .finTableW97          = ");
            builder.Append(" (").Append(GetFinTableW97()).Append(" )\n");
            builder.Append("    .fTtp                 = ");
            builder.Append(" (").Append(GetFTtp()).Append(" )\n");
            builder.Append("    .dxaAbs               = ");
            builder.Append(" (").Append(GetDxaAbs()).Append(" )\n");
            builder.Append("    .dyaAbs               = ");
            builder.Append(" (").Append(GetDyaAbs()).Append(" )\n");
            builder.Append("    .dxaWidth             = ");
            builder.Append(" (").Append(GetDxaWidth()).Append(" )\n");
            builder.Append("    .fBrLnAbove           = ");
            builder.Append(" (").Append(GetFBrLnAbove()).Append(" )\n");
            builder.Append("    .fBrLnBelow           = ");
            builder.Append(" (").Append(GetFBrLnBelow()).Append(" )\n");
            builder.Append("    .pcVert               = ");
            builder.Append(" (").Append(GetPcVert()).Append(" )\n");
            builder.Append("    .pcHorz               = ");
            builder.Append(" (").Append(GetPcHorz()).Append(" )\n");
            builder.Append("    .wr                   = ");
            builder.Append(" (").Append(GetWr()).Append(" )\n");
            builder.Append("    .fNoAutoHyph          = ");
            builder.Append(" (").Append(GetFNoAutoHyph()).Append(" )\n");
            builder.Append("    .dyaHeight            = ");
            builder.Append(" (").Append(GetDyaHeight()).Append(" )\n");
            builder.Append("    .fMinHeight           = ");
            builder.Append(" (").Append(GetFMinHeight()).Append(" )\n");
            builder.Append("    .dcs                  = ");
            builder.Append(" (").Append(GetDcs()).Append(" )\n");
            builder.Append("    .dyaFromText          = ");
            builder.Append(" (").Append(GetDyaFromText()).Append(" )\n");
            builder.Append("    .dxaFromText          = ");
            builder.Append(" (").Append(GetDxaFromText()).Append(" )\n");
            builder.Append("    .fLocked              = ");
            builder.Append(" (").Append(GetFLocked()).Append(" )\n");
            builder.Append("    .fWidowControl        = ");
            builder.Append(" (").Append(GetFWidowControl()).Append(" )\n");
            builder.Append("    .fKinsoku             = ");
            builder.Append(" (").Append(GetFKinsoku()).Append(" )\n");
            builder.Append("    .fWordWrap            = ");
            builder.Append(" (").Append(GetFWordWrap()).Append(" )\n");
            builder.Append("    .fOverflowPunct       = ");
            builder.Append(" (").Append(GetFOverflowPunct()).Append(" )\n");
            builder.Append("    .fTopLinePunct        = ");
            builder.Append(" (").Append(GetFTopLinePunct()).Append(" )\n");
            builder.Append("    .fAutoSpaceDE         = ");
            builder.Append(" (").Append(GetFAutoSpaceDE()).Append(" )\n");
            builder.Append("    .fAutoSpaceDN         = ");
            builder.Append(" (").Append(GetFAutoSpaceDN()).Append(" )\n");
            builder.Append("    .wAlignFont           = ");
            builder.Append(" (").Append(GetWAlignFont()).Append(" )\n");
            builder.Append("    .fontAlign            = ");
            builder.Append(" (").Append(GetFontAlign()).Append(" )\n");
            builder.Append("         .fVertical                = ").Append(IsFVertical()).Append('\n');
            builder.Append("         .fBackward                = ").Append(IsFBackward()).Append('\n');
            builder.Append("         .fRotateFont              = ").Append(IsFRotateFont()).Append('\n');
            builder.Append("    .lvl                  = ");
            builder.Append(" (").Append(GetLvl()).Append(" )\n");
            builder.Append("    .fBiDi                = ");
            builder.Append(" (").Append(GetFBiDi()).Append(" )\n");
            builder.Append("    .fNumRMIns            = ");
            builder.Append(" (").Append(GetFNumRMIns()).Append(" )\n");
            builder.Append("    .fCrLf                = ");
            builder.Append(" (").Append(GetFCrLf()).Append(" )\n");
            builder.Append("    .fUsePgsuSettings     = ");
            builder.Append(" (").Append(GetFUsePgsuSettings()).Append(" )\n");
            builder.Append("    .fAdjustRight         = ");
            builder.Append(" (").Append(GetFAdjustRight()).Append(" )\n");
            builder.Append("    .itap                 = ");
            builder.Append(" (").Append(GetItap()).Append(" )\n");
            builder.Append("    .fInnerTableCell      = ");
            builder.Append(" (").Append(GetFInnerTableCell()).Append(" )\n");
            builder.Append("    .fOpenTch             = ");
            builder.Append(" (").Append(GetFOpenTch()).Append(" )\n");
            builder.Append("    .fTtpEmbedded         = ");
            builder.Append(" (").Append(GetFTtpEmbedded()).Append(" )\n");
            builder.Append("    .dxcRight             = ");
            builder.Append(" (").Append(GetDxcRight()).Append(" )\n");
            builder.Append("    .dxcLeft              = ");
            builder.Append(" (").Append(GetDxcLeft()).Append(" )\n");
            builder.Append("    .dxcLeft1             = ");
            builder.Append(" (").Append(GetDxcLeft1()).Append(" )\n");
            builder.Append("    .fDyaBeforeAuto       = ");
            builder.Append(" (").Append(GetFDyaBeforeAuto()).Append(" )\n");
            builder.Append("    .fDyaAfterAuto        = ");
            builder.Append(" (").Append(GetFDyaAfterAuto()).Append(" )\n");
            builder.Append("    .dxaRight             = ");
            builder.Append(" (").Append(GetDxaRight()).Append(" )\n");
            builder.Append("    .dxaLeft              = ");
            builder.Append(" (").Append(GetDxaLeft()).Append(" )\n");
            builder.Append("    .dxaLeft1             = ");
            builder.Append(" (").Append(GetDxaLeft1()).Append(" )\n");
            builder.Append("    .jc                   = ");
            builder.Append(" (").Append(GetJc()).Append(" )\n");
            builder.Append("    .fNoAllowOverlap      = ");
            builder.Append(" (").Append(GetFNoAllowOverlap()).Append(" )\n");
            builder.Append("    .brcTop               = ");
            builder.Append(" (").Append(GetBrcTop()).Append(" )\n");
            builder.Append("    .brcLeft              = ");
            builder.Append(" (").Append(GetBrcLeft()).Append(" )\n");
            builder.Append("    .brcBottom            = ");
            builder.Append(" (").Append(GetBrcBottom()).Append(" )\n");
            builder.Append("    .brcRight             = ");
            builder.Append(" (").Append(GetBrcRight()).Append(" )\n");
            builder.Append("    .brcBetween           = ");
            builder.Append(" (").Append(GetBrcBetween()).Append(" )\n");
            builder.Append("    .brcBar               = ");
            builder.Append(" (").Append(GetBrcBar()).Append(" )\n");
            builder.Append("    .shd                  = ");
            builder.Append(" (").Append(GetShd()).Append(" )\n");
            builder.Append("    .anld                 = ");
            builder.Append(" (").Append(GetAnld()).Append(" )\n");
            builder.Append("    .phe                  = ");
            builder.Append(" (").Append(GetPhe()).Append(" )\n");
            builder.Append("    .fPropRMark           = ");
            builder.Append(" (").Append(GetFPropRMark()).Append(" )\n");
            builder.Append("    .ibstPropRMark        = ");
            builder.Append(" (").Append(GetIbstPropRMark()).Append(" )\n");
            builder.Append("    .dttmPropRMark        = ");
            builder.Append(" (").Append(GetDttmPropRMark()).Append(" )\n");
            builder.Append("    .itbdMac              = ");
            builder.Append(" (").Append(GetItbdMac()).Append(" )\n");
            builder.Append("    .rgdxaTab             = ");
            builder.Append(" (").Append(GetRgdxaTab()).Append(" )\n");
            builder.Append("    .rgtbd                = ");
            builder.Append(" (").Append(GetRgtbd()).Append(" )\n");
            builder.Append("    .numrm                = ");
            builder.Append(" (").Append(GetNumrm()).Append(" )\n");
            builder.Append("    .ptap                 = ");
            builder.Append(" (").Append(GetPtap()).Append(" )\n");

            builder.Append("[/PAP]\n");
            return builder.ToString();
        }

        /**
         * Index to style descriptor.
         */
        public int GetIstd()
        {
            return field_1_istd;
        }

        /**
         * Index to style descriptor.
         */
        public void SetIstd(int field_1_istd)
        {
            this.field_1_istd = field_1_istd;
        }

        /**
         * Get the fSideBySide field for the PAP record.
         */
        public bool GetFSideBySide()
        {
            return field_2_fSideBySide;
        }

        /**
         * Set the fSideBySide field for the PAP record.
         */
        public void SetFSideBySide(bool field_2_fSideBySide)
        {
            this.field_2_fSideBySide = field_2_fSideBySide;
        }

        /**
         * Get the fKeep field for the PAP record.
         */
        public bool GetFKeep()
        {
            return field_3_fKeep;
        }

        /**
         * Set the fKeep field for the PAP record.
         */
        public void SetFKeep(bool field_3_fKeep)
        {
            this.field_3_fKeep = field_3_fKeep;
        }

        /**
         * Get the fKeepFollow field for the PAP record.
         */
        public bool GetFKeepFollow()
        {
            return field_4_fKeepFollow;
        }

        /**
         * Set the fKeepFollow field for the PAP record.
         */
        public void SetFKeepFollow(bool field_4_fKeepFollow)
        {
            this.field_4_fKeepFollow = field_4_fKeepFollow;
        }

        /**
         * Get the fPageBreakBefore field for the PAP record.
         */
        public bool GetFPageBreakBefore()
        {
            return field_5_fPageBreakBefore;
        }

        /**
         * Set the fPageBreakBefore field for the PAP record.
         */
        public void SetFPageBreakBefore(bool field_5_fPageBreakBefore)
        {
            this.field_5_fPageBreakBefore = field_5_fPageBreakBefore;
        }

        /**
         * Border line style.
         *
         * @return One of 
         * <li>{@link #BRCL_SINGLE}
         * <li>{@link #BRCL_THICK}
         * <li>{@link #BRCL_DOUBLE}
         * <li>{@link #BRCL_SHADOW}
         */
        public byte GetBrcl()
        {
            return field_6_brcl;
        }

        /**
         * Border line style.
         *
         * @param field_6_brcl
         *        One of 
         * <li>{@link #BRCL_SINGLE}
         * <li>{@link #BRCL_THICK}
         * <li>{@link #BRCL_DOUBLE}
         * <li>{@link #BRCL_SHADOW}
         */
        public void SetBrcl(byte field_6_brcl)
        {
            this.field_6_brcl = field_6_brcl;
        }

        /**
         * Rectangle border codes.
         *
         * @return One of 
         * <li>{@link #BRCP_NONE}
         * <li>{@link #BRCP_BORDER_ABOVE}
         * <li>{@link #BRCP_BORDER_BELOW}
         * <li>{@link #BRCP_BOX_AROUND}
         * <li>{@link #BRCP_BAR_TO_LEFT_OF_PARAGRAPH}
         */
        public byte GetBrcp()
        {
            return field_7_brcp;
        }

        /**
         * Rectangle border codes.
         *
         * @param field_7_brcp
         *        One of 
         * <li>{@link #BRCP_NONE}
         * <li>{@link #BRCP_BORDER_ABOVE}
         * <li>{@link #BRCP_BORDER_BELOW}
         * <li>{@link #BRCP_BOX_AROUND}
         * <li>{@link #BRCP_BAR_TO_LEFT_OF_PARAGRAPH}
         */
        public void SetBrcp(byte field_7_brcp)
        {
            this.field_7_brcp = field_7_brcp;
        }

        /**
         * List level if non-zero.
         */
        public byte GetIlvl()
        {
            return field_8_ilvl;
        }

        /**
         * List level if non-zero.
         */
        public void SetIlvl(byte field_8_ilvl)
        {
            this.field_8_ilvl = field_8_ilvl;
        }

        /**
         * 1-based index into the pllfo (lists structure), if non-zero.
         */
        public int GetIlfo()
        {
            return field_9_ilfo;
        }

        /**
         * 1-based index into the pllfo (lists structure), if non-zero.
         */
        public void SetIlfo(int field_9_ilfo)
        {
            this.field_9_ilfo = field_9_ilfo;
        }

        /**
         * No line numbering.
         */
        public bool GetFNoLnn()
        {
            return field_10_fNoLnn;
        }

        /**
         * No line numbering.
         */
        public void SetFNoLnn(bool field_10_fNoLnn)
        {
            this.field_10_fNoLnn = field_10_fNoLnn;
        }

        /**
         * Line spacing descriptor.
         */
        public LineSpacingDescriptor GetLspd()
        {
            return field_11_lspd;
        }

        /**
         * Line spacing descriptor.
         */
        public void SetLspd(LineSpacingDescriptor field_11_lspd)
        {
            this.field_11_lspd = field_11_lspd;
        }

        /**
         * Space before paragraph.
         */
        public int GetDyaBefore()
        {
            return field_12_dyaBefore;
        }

        /**
         * Space before paragraph.
         */
        public void SetDyaBefore(int field_12_dyaBefore)
        {
            this.field_12_dyaBefore = field_12_dyaBefore;
        }

        /**
         * Space after paragraph.
         */
        public int GetDyaAfter()
        {
            return field_13_dyaAfter;
        }

        /**
         * Space after paragraph.
         */
        public void SetDyaAfter(int field_13_dyaAfter)
        {
            this.field_13_dyaAfter = field_13_dyaAfter;
        }

        /**
         * Paragraph is in table flag.
         */
        public bool GetFInTable()
        {
            return field_14_fInTable;
        }

        /**
         * Paragraph is in table flag.
         */
        public void SetFInTable(bool field_14_fInTable)
        {
            this.field_14_fInTable = field_14_fInTable;
        }

        /**
         * Archaic paragraph is in table flag.
         */
        public bool GetFinTableW97()
        {
            return field_15_finTableW97;
        }

        /**
         * Archaic paragraph is in table flag.
         */
        public void SetFinTableW97(bool field_15_finTableW97)
        {
            this.field_15_finTableW97 = field_15_finTableW97;
        }

        /**
         * Table trailer paragraph (last in table row).
         */
        public bool GetFTtp()
        {
            return field_16_fTtp;
        }

        /**
         * Table trailer paragraph (last in table row).
         */
        public void SetFTtp(bool field_16_fTtp)
        {
            this.field_16_fTtp = field_16_fTtp;
        }

        /**
         * Get the dxaAbs field for the PAP record.
         */
        public int GetDxaAbs()
        {
            return field_17_dxaAbs;
        }

        /**
         * Set the dxaAbs field for the PAP record.
         */
        public void SetDxaAbs(int field_17_dxaAbs)
        {
            this.field_17_dxaAbs = field_17_dxaAbs;
        }

        /**
         * Get the dyaAbs field for the PAP record.
         */
        public int GetDyaAbs()
        {
            return field_18_dyaAbs;
        }

        /**
         * Set the dyaAbs field for the PAP record.
         */
        public void SetDyaAbs(int field_18_dyaAbs)
        {
            this.field_18_dyaAbs = field_18_dyaAbs;
        }

        /**
         * Get the dxaWidth field for the PAP record.
         */
        public int GetDxaWidth()
        {
            return field_19_dxaWidth;
        }

        /**
         * Set the dxaWidth field for the PAP record.
         */
        public void SetDxaWidth(int field_19_dxaWidth)
        {
            this.field_19_dxaWidth = field_19_dxaWidth;
        }

        /**
         * Get the fBrLnAbove field for the PAP record.
         */
        public bool GetFBrLnAbove()
        {
            return field_20_fBrLnAbove;
        }

        /**
         * Set the fBrLnAbove field for the PAP record.
         */
        public void SetFBrLnAbove(bool field_20_fBrLnAbove)
        {
            this.field_20_fBrLnAbove = field_20_fBrLnAbove;
        }

        /**
         * Get the fBrLnBelow field for the PAP record.
         */
        public bool GetFBrLnBelow()
        {
            return field_21_fBrLnBelow;
        }

        /**
         * Set the fBrLnBelow field for the PAP record.
         */
        public void SetFBrLnBelow(bool field_21_fBrLnBelow)
        {
            this.field_21_fBrLnBelow = field_21_fBrLnBelow;
        }

        /**
         * Get the pcVert field for the PAP record.
         */
        public byte GetPcVert()
        {
            return field_22_pcVert;
        }

        /**
         * Set the pcVert field for the PAP record.
         */
        public void SetPcVert(byte field_22_pcVert)
        {
            this.field_22_pcVert = field_22_pcVert;
        }

        /**
         * Get the pcHorz field for the PAP record.
         */
        public byte GetPcHorz()
        {
            return field_23_pcHorz;
        }

        /**
         * Set the pcHorz field for the PAP record.
         */
        public void SetPcHorz(byte field_23_pcHorz)
        {
            this.field_23_pcHorz = field_23_pcHorz;
        }

        /**
         * Get the wr field for the PAP record.
         */
        public byte GetWr()
        {
            return field_24_wr;
        }

        /**
         * Set the wr field for the PAP record.
         */
        public void SetWr(byte field_24_wr)
        {
            this.field_24_wr = field_24_wr;
        }

        /**
         * Get the fNoAutoHyph field for the PAP record.
         */
        public bool GetFNoAutoHyph()
        {
            return field_25_fNoAutoHyph;
        }

        /**
         * Set the fNoAutoHyph field for the PAP record.
         */
        public void SetFNoAutoHyph(bool field_25_fNoAutoHyph)
        {
            this.field_25_fNoAutoHyph = field_25_fNoAutoHyph;
        }

        /**
         * Get the dyaHeight field for the PAP record.
         */
        public int GetDyaHeight()
        {
            return field_26_dyaHeight;
        }

        /**
         * Set the dyaHeight field for the PAP record.
         */
        public void SetDyaHeight(int field_26_dyaHeight)
        {
            this.field_26_dyaHeight = field_26_dyaHeight;
        }

        /**
         * Minimum height is exact or auto.
         *
         * @return One of 
         * <li>{@link #FMINHEIGHT_EXACT}
         * <li>{@link #FMINHEIGHT_AT_LEAST}
         */
        public bool GetFMinHeight()
        {
            return field_27_fMinHeight;
        }

        /**
         * Minimum height is exact or auto.
         *
         * @param field_27_fMinHeight
         *        One of 
         * <li>{@link #FMINHEIGHT_EXACT}
         * <li>{@link #FMINHEIGHT_AT_LEAST}
         */
        public void SetFMinHeight(bool field_27_fMinHeight)
        {
            this.field_27_fMinHeight = field_27_fMinHeight;
        }

        /**
         * Get the dcs field for the PAP record.
         */
        public DropCapSpecifier GetDcs()
        {
            return field_28_dcs;
        }

        /**
         * Set the dcs field for the PAP record.
         */
        public void SetDcs(DropCapSpecifier field_28_dcs)
        {
            this.field_28_dcs = field_28_dcs;
        }

        /**
         * Vertical distance between text and absolutely positioned object.
         */
        public int GetDyaFromText()
        {
            return field_29_dyaFromText;
        }

        /**
         * Vertical distance between text and absolutely positioned object.
         */
        public void SetDyaFromText(int field_29_dyaFromText)
        {
            this.field_29_dyaFromText = field_29_dyaFromText;
        }

        /**
         * Horizontal distance between text and absolutely positioned object.
         */
        public int GetDxaFromText()
        {
            return field_30_dxaFromText;
        }

        /**
         * Horizontal distance between text and absolutely positioned object.
         */
        public void SetDxaFromText(int field_30_dxaFromText)
        {
            this.field_30_dxaFromText = field_30_dxaFromText;
        }

        /**
         * Anchor of an absolutely positioned frame is locked.
         */
        public bool GetFLocked()
        {
            return field_31_fLocked;
        }

        /**
         * Anchor of an absolutely positioned frame is locked.
         */
        public void SetFLocked(bool field_31_fLocked)
        {
            this.field_31_fLocked = field_31_fLocked;
        }

        /**
         * 1, Word will prevent widowed lines in this paragraph from being placed at the beginning of a page.
         */
        public bool GetFWidowControl()
        {
            return field_32_fWidowControl;
        }

        /**
         * 1, Word will prevent widowed lines in this paragraph from being placed at the beginning of a page.
         */
        public void SetFWidowControl(bool field_32_fWidowControl)
        {
            this.field_32_fWidowControl = field_32_fWidowControl;
        }

        /**
         * apply Kinsoku rules when performing line wrapping.
         */
        public bool GetFKinsoku()
        {
            return field_33_fKinsoku;
        }

        /**
         * apply Kinsoku rules when performing line wrapping.
         */
        public void SetFKinsoku(bool field_33_fKinsoku)
        {
            this.field_33_fKinsoku = field_33_fKinsoku;
        }

        /**
         * perform word wrap.
         */
        public bool GetFWordWrap()
        {
            return field_34_fWordWrap;
        }

        /**
         * perform word wrap.
         */
        public void SetFWordWrap(bool field_34_fWordWrap)
        {
            this.field_34_fWordWrap = field_34_fWordWrap;
        }

        /**
         * apply overflow punctuation rules when performing line wrapping.
         */
        public bool GetFOverflowPunct()
        {
            return field_35_fOverflowPunct;
        }

        /**
         * apply overflow punctuation rules when performing line wrapping.
         */
        public void SetFOverflowPunct(bool field_35_fOverflowPunct)
        {
            this.field_35_fOverflowPunct = field_35_fOverflowPunct;
        }

        /**
         * perform top line punctuation Processing.
         */
        public bool GetFTopLinePunct()
        {
            return field_36_fTopLinePunct;
        }

        /**
         * perform top line punctuation Processing.
         */
        public void SetFTopLinePunct(bool field_36_fTopLinePunct)
        {
            this.field_36_fTopLinePunct = field_36_fTopLinePunct;
        }

        /**
         * auto space East Asian and alphabetic characters.
         */
        public bool GetFAutoSpaceDE()
        {
            return field_37_fAutoSpaceDE;
        }

        /**
         * auto space East Asian and alphabetic characters.
         */
        public void SetFAutoSpaceDE(bool field_37_fAutoSpaceDE)
        {
            this.field_37_fAutoSpaceDE = field_37_fAutoSpaceDE;
        }

        /**
         * auto space East Asian and numeric characters.
         */
        public bool GetFAutoSpaceDN()
        {
            return field_38_fAutoSpaceDN;
        }

        /**
         * auto space East Asian and numeric characters.
         */
        public void SetFAutoSpaceDN(bool field_38_fAutoSpaceDN)
        {
            this.field_38_fAutoSpaceDN = field_38_fAutoSpaceDN;
        }

        /**
         * Get the wAlignFont field for the PAP record.
         *
         * @return One of 
         * <li>{@link #WALIGNFONT_HANGING}
         * <li>{@link #WALIGNFONT_CENTERED}
         * <li>{@link #WALIGNFONT_ROMAN}
         * <li>{@link #WALIGNFONT_VARIABLE}
         * <li>{@link #WALIGNFONT_AUTO}
         */
        public int GetWAlignFont()
        {
            return field_39_wAlignFont;
        }

        /**
         * Set the wAlignFont field for the PAP record.
         *
         * @param field_39_wAlignFont
         *        One of 
         * <li>{@link #WALIGNFONT_HANGING}
         * <li>{@link #WALIGNFONT_CENTERED}
         * <li>{@link #WALIGNFONT_ROMAN}
         * <li>{@link #WALIGNFONT_VARIABLE}
         * <li>{@link #WALIGNFONT_AUTO}
         */
        public void SetWAlignFont(int field_39_wAlignFont)
        {
            this.field_39_wAlignFont = field_39_wAlignFont;
        }

        /**
         * Used internally by Word.
         */
        public short GetFontAlign()
        {
            return field_40_fontAlign;
        }

        /**
         * Used internally by Word.
         */
        public void SetFontAlign(short field_40_fontAlign)
        {
            this.field_40_fontAlign = field_40_fontAlign;
        }

        /**
         * Outline level.
         */
        public byte GetLvl()
        {
            return field_41_lvl;
        }

        /**
         * Outline level.
         */
        public void SetLvl(byte field_41_lvl)
        {
            this.field_41_lvl = field_41_lvl;
        }

        /**
         * Get the fBiDi field for the PAP record.
         */
        public bool GetFBiDi()
        {
            return field_42_fBiDi;
        }

        /**
         * Set the fBiDi field for the PAP record.
         */
        public void SetFBiDi(bool field_42_fBiDi)
        {
            this.field_42_fBiDi = field_42_fBiDi;
        }

        /**
         * Get the fNumRMIns field for the PAP record.
         */
        public bool GetFNumRMIns()
        {
            return field_43_fNumRMIns;
        }

        /**
         * Set the fNumRMIns field for the PAP record.
         */
        public void SetFNumRMIns(bool field_43_fNumRMIns)
        {
            this.field_43_fNumRMIns = field_43_fNumRMIns;
        }

        /**
         * Get the fCrLf field for the PAP record.
         */
        public bool GetFCrLf()
        {
            return field_44_fCrLf;
        }

        /**
         * Set the fCrLf field for the PAP record.
         */
        public void SetFCrLf(bool field_44_fCrLf)
        {
            this.field_44_fCrLf = field_44_fCrLf;
        }

        /**
         * Get the fUsePgsuSettings field for the PAP record.
         */
        public bool GetFUsePgsuSettings()
        {
            return field_45_fUsePgsuSettings;
        }

        /**
         * Set the fUsePgsuSettings field for the PAP record.
         */
        public void SetFUsePgsuSettings(bool field_45_fUsePgsuSettings)
        {
            this.field_45_fUsePgsuSettings = field_45_fUsePgsuSettings;
        }

        /**
         * Get the fAdjustRight field for the PAP record.
         */
        public bool GetFAdjustRight()
        {
            return field_46_fAdjustRight;
        }

        /**
         * Set the fAdjustRight field for the PAP record.
         */
        public void SetFAdjustRight(bool field_46_fAdjustRight)
        {
            this.field_46_fAdjustRight = field_46_fAdjustRight;
        }

        /**
         * Table nesting level.
         */
        public int GetItap()
        {
            return field_47_itap;
        }

        /**
         * Table nesting level.
         */
        public void SetItap(int field_47_itap)
        {
            this.field_47_itap = field_47_itap;
        }

        /**
         * When 1, the end of paragraph mark is really an end of cell mark for a nested table cell.
         */
        public bool GetFInnerTableCell()
        {
            return field_48_fInnerTableCell;
        }

        /**
         * When 1, the end of paragraph mark is really an end of cell mark for a nested table cell.
         */
        public void SetFInnerTableCell(bool field_48_fInnerTableCell)
        {
            this.field_48_fInnerTableCell = field_48_fInnerTableCell;
        }

        /**
         * Ensure the Table Cell char doesn't show up as zero height.
         */
        public bool GetFOpenTch()
        {
            return field_49_fOpenTch;
        }

        /**
         * Ensure the Table Cell char doesn't show up as zero height.
         */
        public void SetFOpenTch(bool field_49_fOpenTch)
        {
            this.field_49_fOpenTch = field_49_fOpenTch;
        }

        /**
         * Word 97 compatibility indicates this end of paragraph mark is really an end of row marker for a nested table.
         */
        public bool GetFTtpEmbedded()
        {
            return field_50_fTtpEmbedded;
        }

        /**
         * Word 97 compatibility indicates this end of paragraph mark is really an end of row marker for a nested table.
         */
        public void SetFTtpEmbedded(bool field_50_fTtpEmbedded)
        {
            this.field_50_fTtpEmbedded = field_50_fTtpEmbedded;
        }

        /**
         * Right indent in character units.
         */
        public short GetDxcRight()
        {
            return field_51_dxcRight;
        }

        /**
         * Right indent in character units.
         */
        public void SetDxcRight(short field_51_dxcRight)
        {
            this.field_51_dxcRight = field_51_dxcRight;
        }

        /**
         * Left indent in character units.
         */
        public short GetDxcLeft()
        {
            return field_52_dxcLeft;
        }

        /**
         * Left indent in character units.
         */
        public void SetDxcLeft(short field_52_dxcLeft)
        {
            this.field_52_dxcLeft = field_52_dxcLeft;
        }

        /**
         * First line indent in character units.
         */
        public short GetDxcLeft1()
        {
            return field_53_dxcLeft1;
        }

        /**
         * First line indent in character units.
         */
        public void SetDxcLeft1(short field_53_dxcLeft1)
        {
            this.field_53_dxcLeft1 = field_53_dxcLeft1;
        }

        /**
         * Vertical spacing before is automatic.
         */
        public bool GetFDyaBeforeAuto()
        {
            return field_54_fDyaBeforeAuto;
        }

        /**
         * Vertical spacing before is automatic.
         */
        public void SetFDyaBeforeAuto(bool field_54_fDyaBeforeAuto)
        {
            this.field_54_fDyaBeforeAuto = field_54_fDyaBeforeAuto;
        }

        /**
         * Vertical spacing after is automatic.
         */
        public bool GetFDyaAfterAuto()
        {
            return field_55_fDyaAfterAuto;
        }

        /**
         * Vertical spacing after is automatic.
         */
        public void SetFDyaAfterAuto(bool field_55_fDyaAfterAuto)
        {
            this.field_55_fDyaAfterAuto = field_55_fDyaAfterAuto;
        }

        /**
         * Get the dxaRight field for the PAP record.
         */
        public int GetDxaRight()
        {
            return field_56_dxaRight;
        }

        /**
         * Set the dxaRight field for the PAP record.
         */
        public void SetDxaRight(int field_56_dxaRight)
        {
            this.field_56_dxaRight = field_56_dxaRight;
        }

        /**
         * Get the dxaLeft field for the PAP record.
         */
        public int GetDxaLeft()
        {
            return field_57_dxaLeft;
        }

        /**
         * Set the dxaLeft field for the PAP record.
         */
        public void SetDxaLeft(int field_57_dxaLeft)
        {
            this.field_57_dxaLeft = field_57_dxaLeft;
        }

        /**
         * Get the dxaLeft1 field for the PAP record.
         */
        public int GetDxaLeft1()
        {
            return field_58_dxaLeft1;
        }

        /**
         * Set the dxaLeft1 field for the PAP record.
         */
        public void SetDxaLeft1(int field_58_dxaLeft1)
        {
            this.field_58_dxaLeft1 = field_58_dxaLeft1;
        }

        /**
         * Get the jc field for the PAP record.
         */
        public byte GetJc()
        {
            return field_59_jc;
        }

        /**
         * Set the jc field for the PAP record.
         */
        public void SetJc(byte field_59_jc)
        {
            this.field_59_jc = field_59_jc;
        }

        /**
         * Get the fNoAllowOverlap field for the PAP record.
         */
        public bool GetFNoAllowOverlap()
        {
            return field_60_fNoAllowOverlap;
        }

        /**
         * Set the fNoAllowOverlap field for the PAP record.
         */
        public void SetFNoAllowOverlap(bool field_60_fNoAllowOverlap)
        {
            this.field_60_fNoAllowOverlap = field_60_fNoAllowOverlap;
        }

        /**
         * Get the brcTop field for the PAP record.
         */
        public BorderCode GetBrcTop()
        {
            return field_61_brcTop;
        }

        /**
         * Set the brcTop field for the PAP record.
         */
        public void SetBrcTop(BorderCode field_61_brcTop)
        {
            this.field_61_brcTop = field_61_brcTop;
        }

        /**
         * Get the brcLeft field for the PAP record.
         */
        public BorderCode GetBrcLeft()
        {
            return field_62_brcLeft;
        }

        /**
         * Set the brcLeft field for the PAP record.
         */
        public void SetBrcLeft(BorderCode field_62_brcLeft)
        {
            this.field_62_brcLeft = field_62_brcLeft;
        }

        /**
         * Get the brcBottom field for the PAP record.
         */
        public BorderCode GetBrcBottom()
        {
            return field_63_brcBottom;
        }

        /**
         * Set the brcBottom field for the PAP record.
         */
        public void SetBrcBottom(BorderCode field_63_brcBottom)
        {
            this.field_63_brcBottom = field_63_brcBottom;
        }

        /**
         * Get the brcRight field for the PAP record.
         */
        public BorderCode GetBrcRight()
        {
            return field_64_brcRight;
        }

        /**
         * Set the brcRight field for the PAP record.
         */
        public void SetBrcRight(BorderCode field_64_brcRight)
        {
            this.field_64_brcRight = field_64_brcRight;
        }

        /**
         * Get the brcBetween field for the PAP record.
         */
        public BorderCode GetBrcBetween()
        {
            return field_65_brcBetween;
        }

        /**
         * Set the brcBetween field for the PAP record.
         */
        public void SetBrcBetween(BorderCode field_65_brcBetween)
        {
            this.field_65_brcBetween = field_65_brcBetween;
        }

        /**
         * Get the brcBar field for the PAP record.
         */
        public BorderCode GetBrcBar()
        {
            return field_66_brcBar;
        }

        /**
         * Set the brcBar field for the PAP record.
         */
        public void SetBrcBar(BorderCode field_66_brcBar)
        {
            this.field_66_brcBar = field_66_brcBar;
        }

        /**
         * Get the shd field for the PAP record.
         */
        public ShadingDescriptor GetShd()
        {
            return field_67_shd;
        }

        /**
         * Set the shd field for the PAP record.
         */
        public void SetShd(ShadingDescriptor field_67_shd)
        {
            this.field_67_shd = field_67_shd;
        }

        /**
         * Get the anld field for the PAP record.
         */
        public byte[] GetAnld()
        {
            return field_68_anld;
        }

        /**
         * Set the anld field for the PAP record.
         */
        public void SetAnld(byte[] field_68_anld)
        {
            this.field_68_anld = field_68_anld;
        }

        /**
         * Get the phe field for the PAP record.
         */
        public byte[] GetPhe()
        {
            return field_69_phe;
        }

        /**
         * Set the phe field for the PAP record.
         */
        public void SetPhe(byte[] field_69_phe)
        {
            this.field_69_phe = field_69_phe;
        }

        /**
         * Get the fPropRMark field for the PAP record.
         */
        public bool GetFPropRMark()
        {
            return field_70_fPropRMark;
        }

        /**
         * Set the fPropRMark field for the PAP record.
         */
        public void SetFPropRMark(bool field_70_fPropRMark)
        {
            this.field_70_fPropRMark = field_70_fPropRMark;
        }

        /**
         * Get the ibstPropRMark field for the PAP record.
         */
        public int GetIbstPropRMark()
        {
            return field_71_ibstPropRMark;
        }

        /**
         * Set the ibstPropRMark field for the PAP record.
         */
        public void SetIbstPropRMark(int field_71_ibstPropRMark)
        {
            this.field_71_ibstPropRMark = field_71_ibstPropRMark;
        }

        /**
         * Get the dttmPropRMark field for the PAP record.
         */
        public DateAndTime GetDttmPropRMark()
        {
            return field_72_dttmPropRMark;
        }

        /**
         * Set the dttmPropRMark field for the PAP record.
         */
        public void SetDttmPropRMark(DateAndTime field_72_dttmPropRMark)
        {
            this.field_72_dttmPropRMark = field_72_dttmPropRMark;
        }

        /**
         * Get the itbdMac field for the PAP record.
         */
        public int GetItbdMac()
        {
            return field_73_itbdMac;
        }

        /**
         * Set the itbdMac field for the PAP record.
         */
        public void SetItbdMac(int field_73_itbdMac)
        {
            this.field_73_itbdMac = field_73_itbdMac;
        }

        /**
         * Get the rgdxaTab field for the PAP record.
         */
        public int[] GetRgdxaTab()
        {
            return field_74_rgdxaTab;
        }

        /**
         * Set the rgdxaTab field for the PAP record.
         */
        public void SetRgdxaTab(int[] field_74_rgdxaTab)
        {
            this.field_74_rgdxaTab = field_74_rgdxaTab;
        }

        /**
         * Get the rgtbd field for the PAP record.
         */
        public byte[] GetRgtbd()
        {
            return field_75_rgtbd;
        }

        /**
         * Set the rgtbd field for the PAP record.
         */
        public void SetRgtbd(byte[] field_75_rgtbd)
        {
            this.field_75_rgtbd = field_75_rgtbd;
        }

        /**
         * Get the numrm field for the PAP record.
         */
        public byte[] GetNumrm()
        {
            return field_76_numrm;
        }

        /**
         * Set the numrm field for the PAP record.
         */
        public void SetNumrm(byte[] field_76_numrm)
        {
            this.field_76_numrm = field_76_numrm;
        }

        /**
         * Get the ptap field for the PAP record.
         */
        public byte[] GetPtap()
        {
            return field_77_ptap;
        }

        /**
         * Set the ptap field for the PAP record.
         */
        public void SetPtap(byte[] field_77_ptap)
        {
            this.field_77_ptap = field_77_ptap;
        }

        /**
         * Sets the fVertical field value.
         * 
         */
        public void SetFVertical(bool value)
        {
            field_40_fontAlign = (short)fVertical.SetBoolean(field_40_fontAlign, value);


        }

        /**
         * 
         * @return  the fVertical field value.
         */
        public bool IsFVertical()
        {
            return fVertical.IsSet(field_40_fontAlign);

        }

        /**
         * Sets the fBackward field value.
         * 
         */
        public void SetFBackward(bool value)
        {
            field_40_fontAlign = (short)fBackward.SetBoolean(field_40_fontAlign, value);


        }

        /**
         * 
         * @return  the fBackward field value.
         */
        public bool IsFBackward()
        {
            return fBackward.IsSet(field_40_fontAlign);

        }

        /**
         * Sets the fRotateFont field value.
         * 
         */
        public void SetFRotateFont(bool value)
        {
            field_40_fontAlign = (short)fRotateFont.SetBoolean(field_40_fontAlign, value);


        }

        /**
         * 
         * @return  the fRotateFont field value.
         */
        public bool IsFRotateFont()
        {
            return fRotateFont.IsSet(field_40_fontAlign);

        }

    }
}



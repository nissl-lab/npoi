
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
     * Document Properties.
     * NOTE: This source Is automatically generated please do not modify this file.  Either subclass or
     *       remove the record in src/records/definitions.

     * @author S. Ryan Ackley
     */
    public abstract class DOPAbstractType : BaseObject
    {

        protected byte field_1_formatFlags;
        private static BitField fFacingPages = BitFieldFactory.GetInstance(0x01);
        private static BitField fWidowControl = BitFieldFactory.GetInstance(0x02);
        private static BitField fPMHMainDoc = BitFieldFactory.GetInstance(0x04);
        private static BitField grfSupression = BitFieldFactory.GetInstance(0x18);
        private static BitField fpc = BitFieldFactory.GetInstance(0x60);
        private static BitField unused1 = BitFieldFactory.GetInstance(0x80);
        protected byte field_2_unused2;
        protected short field_3_footnoteInfo;
        private static BitField rncFtn = BitFieldFactory.GetInstance(0x0003);
        private static BitField nFtn = BitFieldFactory.GetInstance(0xfffc);
        protected byte field_4_fOutlineDirtySave;
        protected byte field_5_docinfo;
        private static BitField fOnlyMacPics = BitFieldFactory.GetInstance(0x01);
        private static BitField fOnlyWinPics = BitFieldFactory.GetInstance(0x02);
        private static BitField fLabelDoc = BitFieldFactory.GetInstance(0x04);
        private static BitField fHyphCapitals = BitFieldFactory.GetInstance(0x08);
        private static BitField fAutoHyphen = BitFieldFactory.GetInstance(0x10);
        private static BitField fFormNoFields = BitFieldFactory.GetInstance(0x20);
        private static BitField fLinkStyles = BitFieldFactory.GetInstance(0x40);
        private static BitField fRevMarking = BitFieldFactory.GetInstance(0x80);
        protected byte field_6_docinfo1;
        private static BitField fBackup = BitFieldFactory.GetInstance(0x01);
        private static BitField fExactCWords = BitFieldFactory.GetInstance(0x02);
        private static BitField fPagHidden = BitFieldFactory.GetInstance(0x04);
        private static BitField fPagResults = BitFieldFactory.GetInstance(0x08);
        private static BitField fLockAtn = BitFieldFactory.GetInstance(0x10);
        private static BitField fMirrorMargins = BitFieldFactory.GetInstance(0x20);
        private static BitField unused3 = BitFieldFactory.GetInstance(0x40);
        private static BitField fDfltTrueType = BitFieldFactory.GetInstance(0x80);
        protected byte field_7_docinfo2;
        private static BitField fPagSupressTopSpacing = BitFieldFactory.GetInstance(0x01);
        private static BitField fProtEnabled = BitFieldFactory.GetInstance(0x02);
        private static BitField fDispFormFldSel = BitFieldFactory.GetInstance(0x04);
        private static BitField fRMView = BitFieldFactory.GetInstance(0x08);
        private static BitField fRMPrint = BitFieldFactory.GetInstance(0x10);
        private static BitField unused4 = BitFieldFactory.GetInstance(0x20);
        private static BitField fLockRev = BitFieldFactory.GetInstance(0x40);
        private static BitField fEmbedFonts = BitFieldFactory.GetInstance(0x80);
        protected short field_8_docinfo3;
        private static BitField oldfNoTabForInd = BitFieldFactory.GetInstance(0x0001);
        private static BitField oldfNoSpaceRaiseLower = BitFieldFactory.GetInstance(0x0002);
        private static BitField oldfSuppressSpbfAfterPageBreak = BitFieldFactory.GetInstance(0x0004);
        private static BitField oldfWrapTrailSpaces = BitFieldFactory.GetInstance(0x0008);
        private static BitField oldfMapPrintTextColor = BitFieldFactory.GetInstance(0x0010);
        private static BitField oldfNoColumnBalance = BitFieldFactory.GetInstance(0x0020);
        private static BitField oldfConvMailMergeEsc = BitFieldFactory.GetInstance(0x0040);
        private static BitField oldfSupressTopSpacing = BitFieldFactory.GetInstance(0x0080);
        private static BitField oldfOrigWordTableRules = BitFieldFactory.GetInstance(0x0100);
        private static BitField oldfTransparentMetafiles = BitFieldFactory.GetInstance(0x0200);
        private static BitField oldfShowBreaksInFrames = BitFieldFactory.GetInstance(0x0400);
        private static BitField oldfSwapBordersFacingPgs = BitFieldFactory.GetInstance(0x0800);
        private static BitField unused5 = BitFieldFactory.GetInstance(0xf000);
        protected int field_9_dxaTab;
        protected int field_10_wSpare;
        protected int field_11_dxaHotz;
        protected int field_12_cConsexHypLim;
        protected int field_13_wSpare2;
        protected int field_14_dttmCreated;
        protected int field_15_dttmRevised;
        protected int field_16_dttmLastPrint;
        protected int field_17_nRevision;
        protected int field_18_tmEdited;
        protected int field_19_cWords;
        protected int field_20_cCh;
        protected int field_21_cPg;
        protected int field_22_cParas;
        protected short field_23_Edn;
        private static BitField rncEdn = BitFieldFactory.GetInstance(0x0003);
        private static BitField nEdn = BitFieldFactory.GetInstance(0xfffc);
        protected short field_24_Edn1;
        private static BitField epc = BitFieldFactory.GetInstance(0x0003);
        private static BitField nfcFtnRef1 = BitFieldFactory.GetInstance(0x003c);
        private static BitField nfcEdnRef1 = BitFieldFactory.GetInstance(0x03c0);
        private static BitField fPrintFormData = BitFieldFactory.GetInstance(0x0400);
        private static BitField fSaveFormData = BitFieldFactory.GetInstance(0x0800);
        private static BitField fShadeFormData = BitFieldFactory.GetInstance(0x1000);
        private static BitField fWCFtnEdn = BitFieldFactory.GetInstance(0x8000);
        protected int field_25_cLines;
        protected int field_26_cWordsFtnEnd;
        protected int field_27_cChFtnEdn;
        protected short field_28_cPgFtnEdn;
        protected int field_29_cParasFtnEdn;
        protected int field_30_cLinesFtnEdn;
        protected int field_31_lKeyProtDoc;
        protected short field_32_view;
        private static BitField wvkSaved = BitFieldFactory.GetInstance(0x0007);
        private static BitField wScaleSaved = BitFieldFactory.GetInstance(0x0ff8);
        private static BitField zkSaved = BitFieldFactory.GetInstance(0x3000);
        private static BitField fRotateFontW6 = BitFieldFactory.GetInstance(0x4000);
        private static BitField iGutterPos = BitFieldFactory.GetInstance(0x8000);
        protected int field_33_docinfo4;
        private static BitField fNoTabForInd = BitFieldFactory.GetInstance(0x00000001);
        private static BitField fNoSpaceRaiseLower = BitFieldFactory.GetInstance(0x00000002);
        private static BitField fSupressSpdfAfterPageBreak = BitFieldFactory.GetInstance(0x00000004);
        private static BitField fWrapTrailSpaces = BitFieldFactory.GetInstance(0x00000008);
        private static BitField fMapPrintTextColor = BitFieldFactory.GetInstance(0x00000010);
        private static BitField fNoColumnBalance = BitFieldFactory.GetInstance(0x00000020);
        private static BitField fConvMailMergeEsc = BitFieldFactory.GetInstance(0x00000040);
        private static BitField fSupressTopSpacing = BitFieldFactory.GetInstance(0x00000080);
        private static BitField fOrigWordTableRules = BitFieldFactory.GetInstance(0x00000100);
        private static BitField fTransparentMetafiles = BitFieldFactory.GetInstance(0x00000200);
        private static BitField fShowBreaksInFrames = BitFieldFactory.GetInstance(0x00000400);
        private static BitField fSwapBordersFacingPgs = BitFieldFactory.GetInstance(0x00000800);
        private static BitField fSuppressTopSPacingMac5 = BitFieldFactory.GetInstance(0x00010000);
        private static BitField fTruncDxaExpand = BitFieldFactory.GetInstance(0x00020000);
        private static BitField fPrintBodyBeforeHdr = BitFieldFactory.GetInstance(0x00040000);
        private static BitField fNoLeading = BitFieldFactory.GetInstance(0x00080000);
        private static BitField fMWSmallCaps = BitFieldFactory.GetInstance(0x00200000);
        protected short field_34_adt;
        protected byte[] field_35_doptypography;
        protected byte[] field_36_dogrid;
        protected short field_37_docinfo5;
        private static BitField lvl = BitFieldFactory.GetInstance(0x001e);
        private static BitField fGramAllDone = BitFieldFactory.GetInstance(0x0020);
        private static BitField fGramAllClean = BitFieldFactory.GetInstance(0x0040);
        private static BitField fSubsetFonts = BitFieldFactory.GetInstance(0x0080);
        private static BitField fHideLastVersion = BitFieldFactory.GetInstance(0x0100);
        private static BitField fHtmlDoc = BitFieldFactory.GetInstance(0x0200);
        private static BitField fSnapBorder = BitFieldFactory.GetInstance(0x0800);
        private static BitField fIncludeHeader = BitFieldFactory.GetInstance(0x1000);
        private static BitField fIncludeFooter = BitFieldFactory.GetInstance(0x2000);
        private static BitField fForcePageSizePag = BitFieldFactory.GetInstance(0x4000);
        private static BitField fMinFontSizePag = BitFieldFactory.GetInstance(0x8000);
        protected short field_38_docinfo6;
        private static BitField fHaveVersions = BitFieldFactory.GetInstance(0x0001);
        private static BitField fAutoVersions = BitFieldFactory.GetInstance(0x0002);
        protected byte[] field_39_asumyi;
        protected int field_40_cChWS;
        protected int field_41_cChWSFtnEdn;
        protected int field_42_grfDocEvents;
        protected int field_43_virusinfo;
        private static BitField fVirusPrompted = BitFieldFactory.GetInstance(0x0001);
        private static BitField fVirusLoadSafe = BitFieldFactory.GetInstance(0x0002);
        private static BitField KeyVirusSession30 = BitFieldFactory.GetInstance(unchecked((int)0xfffffffc));
        protected byte[] field_44_Spare;
        protected int field_45_reserved1;
        protected int field_46_reserved2;
        protected int field_47_cDBC;
        protected int field_48_cDBCFtnEdn;
        protected int field_49_reserved;
        protected short field_50_nfcFtnRef;
        protected short field_51_nfcEdnRef;
        protected short field_52_hpsZoonFontPag;
        protected short field_53_dywDispPag;


        public DOPAbstractType()
        {

        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_formatFlags = data[0x0 + offset];
            field_2_unused2 = data[0x1 + offset];
            field_3_footnoteInfo = LittleEndian.GetShort(data, 0x2 + offset);
            field_4_fOutlineDirtySave = data[0x4 + offset];
            field_5_docinfo = data[0x5 + offset];
            field_6_docinfo1 = data[0x6 + offset];
            field_7_docinfo2 = data[0x7 + offset];
            field_8_docinfo3 = LittleEndian.GetShort(data, 0x8 + offset);
            field_9_dxaTab = LittleEndian.GetShort(data, 0xa + offset);
            field_10_wSpare = LittleEndian.GetShort(data, 0xc + offset);
            field_11_dxaHotz = LittleEndian.GetShort(data, 0xe + offset);
            field_12_cConsexHypLim = LittleEndian.GetShort(data, 0x10 + offset);
            field_13_wSpare2 = LittleEndian.GetShort(data, 0x12 + offset);
            field_14_dttmCreated = LittleEndian.GetInt(data, 0x14 + offset);
            field_15_dttmRevised = LittleEndian.GetInt(data, 0x18 + offset);
            field_16_dttmLastPrint = LittleEndian.GetInt(data, 0x1c + offset);
            field_17_nRevision = LittleEndian.GetShort(data, 0x20 + offset);
            field_18_tmEdited = LittleEndian.GetInt(data, 0x22 + offset);
            field_19_cWords = LittleEndian.GetInt(data, 0x26 + offset);
            field_20_cCh = LittleEndian.GetInt(data, 0x2a + offset);
            field_21_cPg = LittleEndian.GetShort(data, 0x2e + offset);
            field_22_cParas = LittleEndian.GetInt(data, 0x30 + offset);
            field_23_Edn = LittleEndian.GetShort(data, 0x34 + offset);
            field_24_Edn1 = LittleEndian.GetShort(data, 0x36 + offset);
            field_25_cLines = LittleEndian.GetInt(data, 0x38 + offset);
            field_26_cWordsFtnEnd = LittleEndian.GetInt(data, 0x3c + offset);
            field_27_cChFtnEdn = LittleEndian.GetInt(data, 0x40 + offset);
            field_28_cPgFtnEdn = LittleEndian.GetShort(data, 0x44 + offset);
            field_29_cParasFtnEdn = LittleEndian.GetInt(data, 0x46 + offset);
            field_30_cLinesFtnEdn = LittleEndian.GetInt(data, 0x4a + offset);
            field_31_lKeyProtDoc = LittleEndian.GetInt(data, 0x4e + offset);
            field_32_view = LittleEndian.GetShort(data, 0x52 + offset);
            field_33_docinfo4 = LittleEndian.GetInt(data, 0x54 + offset);
            field_34_adt = LittleEndian.GetShort(data, 0x58 + offset);
            field_35_doptypography = LittleEndian.GetByteArray(data, 0x5a + offset, 310);
            field_36_dogrid = LittleEndian.GetByteArray(data, 0x190 + offset, 10);
            field_37_docinfo5 = LittleEndian.GetShort(data, 0x19a + offset);
            field_38_docinfo6 = LittleEndian.GetShort(data, 0x19c + offset);
            field_39_asumyi = LittleEndian.GetByteArray(data, 0x19e + offset, 12);
            field_40_cChWS = LittleEndian.GetInt(data, 0x1aa + offset);
            field_41_cChWSFtnEdn = LittleEndian.GetInt(data, 0x1ae + offset);
            field_42_grfDocEvents = LittleEndian.GetInt(data, 0x1b2 + offset);
            field_43_virusinfo = LittleEndian.GetInt(data, 0x1b6 + offset);
            field_44_Spare = LittleEndian.GetByteArray(data, 0x1ba + offset, 30);
            field_45_reserved1 = LittleEndian.GetInt(data, 0x1d8 + offset);
            field_46_reserved2 = LittleEndian.GetInt(data, 0x1dc + offset);
            field_47_cDBC = LittleEndian.GetInt(data, 0x1e0 + offset);
            field_48_cDBCFtnEdn = LittleEndian.GetInt(data, 0x1e4 + offset);
            field_49_reserved = LittleEndian.GetInt(data, 0x1e8 + offset);
            field_50_nfcFtnRef = LittleEndian.GetShort(data, 0x1ec + offset);
            field_51_nfcEdnRef = LittleEndian.GetShort(data, 0x1ee + offset);
            field_52_hpsZoonFontPag = LittleEndian.GetShort(data, 0x1f0 + offset);
            field_53_dywDispPag = LittleEndian.GetShort(data, 0x1f2 + offset);

        }

        public void Serialize(byte[] data, int offset)
        {
            data[0x0 + offset] = field_1_formatFlags; ;
            data[0x1 + offset] = field_2_unused2; ;
            LittleEndian.PutShort(data, 0x2 + offset, (short)field_3_footnoteInfo); ;
            data[0x4 + offset] = field_4_fOutlineDirtySave; ;
            data[0x5 + offset] = field_5_docinfo; ;
            data[0x6 + offset] = field_6_docinfo1; ;
            data[0x7 + offset] = field_7_docinfo2; ;
            LittleEndian.PutShort(data, 0x8 + offset, (short)field_8_docinfo3); ;
            LittleEndian.PutShort(data, 0xa + offset, (short)field_9_dxaTab); ;
            LittleEndian.PutShort(data, 0xc + offset, (short)field_10_wSpare); ;
            LittleEndian.PutShort(data, 0xe + offset, (short)field_11_dxaHotz); ;
            LittleEndian.PutShort(data, 0x10 + offset, (short)field_12_cConsexHypLim); ;
            LittleEndian.PutShort(data, 0x12 + offset, (short)field_13_wSpare2); ;
            LittleEndian.PutInt(data, 0x14 + offset, field_14_dttmCreated); ;
            LittleEndian.PutInt(data, 0x18 + offset, field_15_dttmRevised); ;
            LittleEndian.PutInt(data, 0x1c + offset, field_16_dttmLastPrint); ;
            LittleEndian.PutShort(data, 0x20 + offset, (short)field_17_nRevision); ;
            LittleEndian.PutInt(data, 0x22 + offset, field_18_tmEdited); ;
            LittleEndian.PutInt(data, 0x26 + offset, field_19_cWords); ;
            LittleEndian.PutInt(data, 0x2a + offset, field_20_cCh); ;
            LittleEndian.PutShort(data, 0x2e + offset, (short)field_21_cPg); ;
            LittleEndian.PutInt(data, 0x30 + offset, field_22_cParas); ;
            LittleEndian.PutShort(data, 0x34 + offset, (short)field_23_Edn); ;
            LittleEndian.PutShort(data, 0x36 + offset, (short)field_24_Edn1); ;
            LittleEndian.PutInt(data, 0x38 + offset, field_25_cLines); ;
            LittleEndian.PutInt(data, 0x3c + offset, field_26_cWordsFtnEnd); ;
            LittleEndian.PutInt(data, 0x40 + offset, field_27_cChFtnEdn); ;
            LittleEndian.PutShort(data, 0x44 + offset, (short)field_28_cPgFtnEdn); ;
            LittleEndian.PutInt(data, 0x46 + offset, field_29_cParasFtnEdn); ;
            LittleEndian.PutInt(data, 0x4a + offset, field_30_cLinesFtnEdn); ;
            LittleEndian.PutInt(data, 0x4e + offset, field_31_lKeyProtDoc); ;
            LittleEndian.PutShort(data, 0x52 + offset, (short)field_32_view); ;
            LittleEndian.PutInt(data, 0x54 + offset, field_33_docinfo4); ;
            LittleEndian.PutShort(data, 0x58 + offset, (short)field_34_adt); ;
            Array.Copy(field_35_doptypography, 0, data, 0x5a + offset, field_35_doptypography.Length); ;
            Array.Copy(field_36_dogrid, 0, data, 0x190 + offset, field_36_dogrid.Length); ;
            LittleEndian.PutShort(data, 0x19a + offset, (short)field_37_docinfo5); ;
            LittleEndian.PutShort(data, 0x19c + offset, (short)field_38_docinfo6); ;
            Array.Copy(field_39_asumyi, 0, data, 0x19e + offset, field_39_asumyi.Length); ;
            LittleEndian.PutInt(data, 0x1aa + offset, field_40_cChWS); ;
            LittleEndian.PutInt(data, 0x1ae + offset, field_41_cChWSFtnEdn); ;
            LittleEndian.PutInt(data, 0x1b2 + offset, field_42_grfDocEvents); ;
            LittleEndian.PutInt(data, 0x1b6 + offset, field_43_virusinfo); ;
            Array.Copy(field_44_Spare, 0, data, 0x1ba + offset, field_44_Spare.Length); ;
            LittleEndian.PutInt(data, 0x1d8 + offset, field_45_reserved1); ;
            LittleEndian.PutInt(data, 0x1dc + offset, field_46_reserved2); ;
            LittleEndian.PutInt(data, 0x1e0 + offset, field_47_cDBC); ;
            LittleEndian.PutInt(data, 0x1e4 + offset, field_48_cDBCFtnEdn); ;
            LittleEndian.PutInt(data, 0x1e8 + offset, field_49_reserved); ;
            LittleEndian.PutShort(data, 0x1ec + offset, (short)field_50_nfcFtnRef); ;
            LittleEndian.PutShort(data, 0x1ee + offset, (short)field_51_nfcEdnRef); ;
            LittleEndian.PutShort(data, 0x1f0 + offset, (short)field_52_hpsZoonFontPag); ;
            LittleEndian.PutShort(data, 0x1f2 + offset, (short)field_53_dywDispPag); ;

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DOP]\n");

            buffer.Append("    .formatFlags          = ");
            buffer.Append(" (").Append(GetFormatFlags()).Append(" )\n");
            buffer.Append("         .fFacingPages             = ").Append(IsFFacingPages()).Append('\n');
            buffer.Append("         .fWidowControl            = ").Append(IsFWidowControl()).Append('\n');
            buffer.Append("         .fPMHMainDoc              = ").Append(IsFPMHMainDoc()).Append('\n');
            buffer.Append("         .grfSupression            = ").Append(GetGrfSupression()).Append('\n');
            buffer.Append("         .fpc                      = ").Append(GetFpc()).Append('\n');
            buffer.Append("         .unused1                  = ").Append(IsUnused1()).Append('\n');

            buffer.Append("    .unused2              = ");
            buffer.Append(" (").Append(GetUnused2()).Append(" )\n");

            buffer.Append("    .footnoteInfo         = ");
            buffer.Append(" (").Append(GetFootnoteInfo()).Append(" )\n");
            buffer.Append("         .rncFtn                   = ").Append(GetRncFtn()).Append('\n');
            buffer.Append("         .nFtn                     = ").Append(GetNFtn()).Append('\n');

            buffer.Append("    .fOutlineDirtySave    = ");
            buffer.Append(" (").Append(GetFOutlineDirtySave()).Append(" )\n");

            buffer.Append("    .docinfo              = ");
            buffer.Append(" (").Append(GetDocinfo()).Append(" )\n");
            buffer.Append("         .fOnlyMacPics             = ").Append(IsFOnlyMacPics()).Append('\n');
            buffer.Append("         .fOnlyWinPics             = ").Append(IsFOnlyWinPics()).Append('\n');
            buffer.Append("         .fLabelDoc                = ").Append(IsFLabelDoc()).Append('\n');
            buffer.Append("         .fHyphCapitals            = ").Append(IsFHyphCapitals()).Append('\n');
            buffer.Append("         .fAutoHyphen              = ").Append(IsFAutoHyphen()).Append('\n');
            buffer.Append("         .fFormNoFields            = ").Append(IsFFormNoFields()).Append('\n');
            buffer.Append("         .fLinkStyles              = ").Append(IsFLinkStyles()).Append('\n');
            buffer.Append("         .fRevMarking              = ").Append(IsFRevMarking()).Append('\n');

            buffer.Append("    .docinfo1             = ");
            buffer.Append(" (").Append(GetDocinfo1()).Append(" )\n");
            buffer.Append("         .fBackup                  = ").Append(IsFBackup()).Append('\n');
            buffer.Append("         .fExactCWords             = ").Append(IsFExactCWords()).Append('\n');
            buffer.Append("         .fPagHidden               = ").Append(IsFPagHidden()).Append('\n');
            buffer.Append("         .fPagResults              = ").Append(IsFPagResults()).Append('\n');
            buffer.Append("         .fLockAtn                 = ").Append(IsFLockAtn()).Append('\n');
            buffer.Append("         .fMirrorMargins           = ").Append(IsFMirrorMargins()).Append('\n');
            buffer.Append("         .unused3                  = ").Append(IsUnused3()).Append('\n');
            buffer.Append("         .fDfltTrueType            = ").Append(IsFDfltTrueType()).Append('\n');

            buffer.Append("    .docinfo2             = ");
            buffer.Append(" (").Append(GetDocinfo2()).Append(" )\n");
            buffer.Append("         .fPagSupressTopSpacing     = ").Append(IsFPagSupressTopSpacing()).Append('\n');
            buffer.Append("         .fProtEnabled             = ").Append(IsFProtEnabled()).Append('\n');
            buffer.Append("         .fDispFormFldSel          = ").Append(IsFDispFormFldSel()).Append('\n');
            buffer.Append("         .fRMView                  = ").Append(IsFRMView()).Append('\n');
            buffer.Append("         .fRMPrint                 = ").Append(IsFRMPrint()).Append('\n');
            buffer.Append("         .unused4                  = ").Append(IsUnused4()).Append('\n');
            buffer.Append("         .fLockRev                 = ").Append(IsFLockRev()).Append('\n');
            buffer.Append("         .fEmbedFonts              = ").Append(IsFEmbedFonts()).Append('\n');

            buffer.Append("    .docinfo3             = ");
            buffer.Append(" (").Append(GetDocinfo3()).Append(" )\n");
            buffer.Append("         .oldfNoTabForInd          = ").Append(IsOldfNoTabForInd()).Append('\n');
            buffer.Append("         .oldfNoSpaceRaiseLower     = ").Append(IsOldfNoSpaceRaiseLower()).Append('\n');
            buffer.Append("         .oldfSuppressSpbfAfterPageBreak     = ").Append(IsOldfSuppressSpbfAfterPageBreak()).Append('\n');
            buffer.Append("         .oldfWrapTrailSpaces      = ").Append(IsOldfWrapTrailSpaces()).Append('\n');
            buffer.Append("         .oldfMapPrintTextColor     = ").Append(IsOldfMapPrintTextColor()).Append('\n');
            buffer.Append("         .oldfNoColumnBalance      = ").Append(IsOldfNoColumnBalance()).Append('\n');
            buffer.Append("         .oldfConvMailMergeEsc     = ").Append(IsOldfConvMailMergeEsc()).Append('\n');
            buffer.Append("         .oldfSupressTopSpacing     = ").Append(IsOldfSupressTopSpacing()).Append('\n');
            buffer.Append("         .oldfOrigWordTableRules     = ").Append(IsOldfOrigWordTableRules()).Append('\n');
            buffer.Append("         .oldfTransparentMetafiles     = ").Append(IsOldfTransparentMetafiles()).Append('\n');
            buffer.Append("         .oldfShowBreaksInFrames     = ").Append(IsOldfShowBreaksInFrames()).Append('\n');
            buffer.Append("         .oldfSwapBordersFacingPgs     = ").Append(IsOldfSwapBordersFacingPgs()).Append('\n');
            buffer.Append("         .unused5                  = ").Append(GetUnused5()).Append('\n');

            buffer.Append("    .dxaTab               = ");
            buffer.Append(" (").Append(GetDxaTab()).Append(" )\n");

            buffer.Append("    .wSpare               = ");
            buffer.Append(" (").Append(GetWSpare()).Append(" )\n");

            buffer.Append("    .dxaHotz              = ");
            buffer.Append(" (").Append(GetDxaHotz()).Append(" )\n");

            buffer.Append("    .cConsexHypLim        = ");
            buffer.Append(" (").Append(GetCConsexHypLim()).Append(" )\n");

            buffer.Append("    .wSpare2              = ");
            buffer.Append(" (").Append(GetWSpare2()).Append(" )\n");

            buffer.Append("    .dttmCreated          = ");
            buffer.Append(" (").Append(GetDttmCreated()).Append(" )\n");

            buffer.Append("    .dttmRevised          = ");
            buffer.Append(" (").Append(GetDttmRevised()).Append(" )\n");

            buffer.Append("    .dttmLastPrint        = ");
            buffer.Append(" (").Append(GetDttmLastPrint()).Append(" )\n");

            buffer.Append("    .nRevision            = ");
            buffer.Append(" (").Append(GetNRevision()).Append(" )\n");

            buffer.Append("    .tmEdited             = ");
            buffer.Append(" (").Append(GetTmEdited()).Append(" )\n");

            buffer.Append("    .cWords               = ");
            buffer.Append(" (").Append(GetCWords()).Append(" )\n");

            buffer.Append("    .cCh                  = ");
            buffer.Append(" (").Append(GetCCh()).Append(" )\n");

            buffer.Append("    .cPg                  = ");
            buffer.Append(" (").Append(GetCPg()).Append(" )\n");

            buffer.Append("    .cParas               = ");
            buffer.Append(" (").Append(GetCParas()).Append(" )\n");

            buffer.Append("    .Edn                  = ");
            buffer.Append(" (").Append(GetEdn()).Append(" )\n");
            buffer.Append("         .rncEdn                   = ").Append(GetRncEdn()).Append('\n');
            buffer.Append("         .nEdn                     = ").Append(GetNEdn()).Append('\n');

            buffer.Append("    .Edn1                 = ");
            buffer.Append(" (").Append(GetEdn1()).Append(" )\n");
            buffer.Append("         .epc                      = ").Append(GetEpc()).Append('\n');
            buffer.Append("         .nfcFtnRef1               = ").Append(GetNfcFtnRef1()).Append('\n');
            buffer.Append("         .nfcEdnRef1               = ").Append(GetNfcEdnRef1()).Append('\n');
            buffer.Append("         .fPrintFormData           = ").Append(IsFPrintFormData()).Append('\n');
            buffer.Append("         .fSaveFormData            = ").Append(IsFSaveFormData()).Append('\n');
            buffer.Append("         .fShadeFormData           = ").Append(IsFShadeFormData()).Append('\n');
            buffer.Append("         .fWCFtnEdn                = ").Append(IsFWCFtnEdn()).Append('\n');

            buffer.Append("    .cLines               = ");
            buffer.Append(" (").Append(GetCLines()).Append(" )\n");

            buffer.Append("    .cWordsFtnEnd         = ");
            buffer.Append(" (").Append(GetCWordsFtnEnd()).Append(" )\n");

            buffer.Append("    .cChFtnEdn            = ");
            buffer.Append(" (").Append(GetCChFtnEdn()).Append(" )\n");

            buffer.Append("    .cPgFtnEdn            = ");
            buffer.Append(" (").Append(GetCPgFtnEdn()).Append(" )\n");

            buffer.Append("    .cParasFtnEdn         = ");
            buffer.Append(" (").Append(GetCParasFtnEdn()).Append(" )\n");

            buffer.Append("    .cLinesFtnEdn         = ");
            buffer.Append(" (").Append(GetCLinesFtnEdn()).Append(" )\n");

            buffer.Append("    .lKeyProtDoc          = ");
            buffer.Append(" (").Append(GetLKeyProtDoc()).Append(" )\n");

            buffer.Append("    .view                 = ");
            buffer.Append(" (").Append(GetView()).Append(" )\n");
            buffer.Append("         .wvkSaved                 = ").Append(GetWvkSaved()).Append('\n');
            buffer.Append("         .wScaleSaved              = ").Append(GetWScaleSaved()).Append('\n');
            buffer.Append("         .zkSaved                  = ").Append(GetZkSaved()).Append('\n');
            buffer.Append("         .fRotateFontW6            = ").Append(IsFRotateFontW6()).Append('\n');
            buffer.Append("         .iGutterPos               = ").Append(IsIGutterPos()).Append('\n');

            buffer.Append("    .docinfo4             = ");
            buffer.Append(" (").Append(GetDocinfo4()).Append(" )\n");
            buffer.Append("         .fNoTabForInd             = ").Append(IsFNoTabForInd()).Append('\n');
            buffer.Append("         .fNoSpaceRaiseLower       = ").Append(IsFNoSpaceRaiseLower()).Append('\n');
            buffer.Append("         .fSupressSpdfAfterPageBreak     = ").Append(IsFSupressSpdfAfterPageBreak()).Append('\n');
            buffer.Append("         .fWrapTrailSpaces         = ").Append(IsFWrapTrailSpaces()).Append('\n');
            buffer.Append("         .fMapPrintTextColor       = ").Append(IsFMapPrintTextColor()).Append('\n');
            buffer.Append("         .fNoColumnBalance         = ").Append(IsFNoColumnBalance()).Append('\n');
            buffer.Append("         .fConvMailMergeEsc        = ").Append(IsFConvMailMergeEsc()).Append('\n');
            buffer.Append("         .fSupressTopSpacing       = ").Append(IsFSupressTopSpacing()).Append('\n');
            buffer.Append("         .fOrigWordTableRules      = ").Append(IsFOrigWordTableRules()).Append('\n');
            buffer.Append("         .fTransparentMetafiles     = ").Append(IsFTransparentMetafiles()).Append('\n');
            buffer.Append("         .fShowBreaksInFrames      = ").Append(IsFShowBreaksInFrames()).Append('\n');
            buffer.Append("         .fSwapBordersFacingPgs     = ").Append(IsFSwapBordersFacingPgs()).Append('\n');
            buffer.Append("         .fSuppressTopSPacingMac5     = ").Append(IsFSuppressTopSPacingMac5()).Append('\n');
            buffer.Append("         .fTruncDxaExpand          = ").Append(IsFTruncDxaExpand()).Append('\n');
            buffer.Append("         .fPrintBodyBeforeHdr      = ").Append(IsFPrintBodyBeforeHdr()).Append('\n');
            buffer.Append("         .fNoLeading               = ").Append(IsFNoLeading()).Append('\n');
            buffer.Append("         .fMWSmallCaps             = ").Append(IsFMWSmallCaps()).Append('\n');

            buffer.Append("    .adt                  = ");
            buffer.Append(" (").Append(GetAdt()).Append(" )\n");

            buffer.Append("    .doptypography        = ");
            buffer.Append(" (").Append(GetDoptypography()).Append(" )\n");

            buffer.Append("    .dogrid               = ");
            buffer.Append(" (").Append(GetDogrid()).Append(" )\n");

            buffer.Append("    .docinfo5             = ");
            buffer.Append(" (").Append(GetDocinfo5()).Append(" )\n");
            buffer.Append("         .lvl                      = ").Append(GetLvl()).Append('\n');
            buffer.Append("         .fGramAllDone             = ").Append(IsFGramAllDone()).Append('\n');
            buffer.Append("         .fGramAllClean            = ").Append(IsFGramAllClean()).Append('\n');
            buffer.Append("         .fSubsetFonts             = ").Append(IsFSubsetFonts()).Append('\n');
            buffer.Append("         .fHideLastVersion         = ").Append(IsFHideLastVersion()).Append('\n');
            buffer.Append("         .fHtmlDoc                 = ").Append(IsFHtmlDoc()).Append('\n');
            buffer.Append("         .fSnapBorder              = ").Append(IsFSnapBorder()).Append('\n');
            buffer.Append("         .fIncludeHeader           = ").Append(IsFIncludeHeader()).Append('\n');
            buffer.Append("         .fIncludeFooter           = ").Append(IsFIncludeFooter()).Append('\n');
            buffer.Append("         .fForcePageSizePag        = ").Append(IsFForcePageSizePag()).Append('\n');
            buffer.Append("         .fMinFontSizePag          = ").Append(IsFMinFontSizePag()).Append('\n');

            buffer.Append("    .docinfo6             = ");
            buffer.Append(" (").Append(GetDocinfo6()).Append(" )\n");
            buffer.Append("         .fHaveVersions            = ").Append(IsFHaveVersions()).Append('\n');
            buffer.Append("         .fAutoVersions            = ").Append(IsFAutoVersions()).Append('\n');

            buffer.Append("    .asumyi               = ");
            buffer.Append(" (").Append(GetAsumyi()).Append(" )\n");

            buffer.Append("    .cChWS                = ");
            buffer.Append(" (").Append(GetCChWS()).Append(" )\n");

            buffer.Append("    .cChWSFtnEdn          = ");
            buffer.Append(" (").Append(GetCChWSFtnEdn()).Append(" )\n");

            buffer.Append("    .grfDocEvents         = ");
            buffer.Append(" (").Append(GetGrfDocEvents()).Append(" )\n");

            buffer.Append("    .virusinfo            = ");
            buffer.Append(" (").Append(GetVirusinfo()).Append(" )\n");
            buffer.Append("         .fVirusPrompted           = ").Append(IsFVirusPrompted()).Append('\n');
            buffer.Append("         .fVirusLoadSafe           = ").Append(IsFVirusLoadSafe()).Append('\n');
            buffer.Append("         .KeyVirusSession30        = ").Append(GetKeyVirusSession30()).Append('\n');

            buffer.Append("    .Spare                = ");
            buffer.Append(" (").Append(GetSpare()).Append(" )\n");

            buffer.Append("    .reserved1            = ");
            buffer.Append(" (").Append(GetReserved1()).Append(" )\n");

            buffer.Append("    .reserved2            = ");
            buffer.Append(" (").Append(GetReserved2()).Append(" )\n");

            buffer.Append("    .cDBC                 = ");
            buffer.Append(" (").Append(GetCDBC()).Append(" )\n");

            buffer.Append("    .cDBCFtnEdn           = ");
            buffer.Append(" (").Append(GetCDBCFtnEdn()).Append(" )\n");

            buffer.Append("    .reserved             = ");
            buffer.Append(" (").Append(GetReserved()).Append(" )\n");

            buffer.Append("    .nfcFtnRef            = ");
            buffer.Append(" (").Append(GetNfcFtnRef()).Append(" )\n");

            buffer.Append("    .nfcEdnRef            = ");
            buffer.Append(" (").Append(GetNfcEdnRef()).Append(" )\n");

            buffer.Append("    .hpsZoonFontPag       = ");
            buffer.Append(" (").Append(GetHpsZoonFontPag()).Append(" )\n");

            buffer.Append("    .dywDispPag           = ");
            buffer.Append(" (").Append(GetDywDispPag()).Append(" )\n");

            buffer.Append("[/DOP]\n");
            return buffer.ToString();
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public int GetSize()
        {
            return 4 + +1 + 1 + 2 + 1 + 1 + 1 + 1 + 2 + 2 + 2 + 2 + 2 + 2 + 4 + 4 + 4 + 2 + 4 + 4 + 4 + 2 + 4 + 2 + 2 + 4 + 4 + 4 + 2 + 4 + 4 + 4 + 2 + 4 + 2 + 310 + 10 + 2 + 2 + 12 + 4 + 4 + 4 + 4 + 30 + 4 + 4 + 4 + 4 + 4 + 2 + 2 + 2 + 2;
        }



        /**
         * Get the formatFlags field for the DOP record.
         */
        public byte GetFormatFlags()
        {
            return field_1_formatFlags;
        }

        /**
         * Set the formatFlags field for the DOP record.
         */
        public void SetFormatFlags(byte field_1_formatFlags)
        {
            this.field_1_formatFlags = field_1_formatFlags;
        }

        /**
         * Get the unused2 field for the DOP record.
         */
        public byte GetUnused2()
        {
            return field_2_unused2;
        }

        /**
         * Set the unused2 field for the DOP record.
         */
        public void SetUnused2(byte field_2_unused2)
        {
            this.field_2_unused2 = field_2_unused2;
        }

        /**
         * Get the footnoteInfo field for the DOP record.
         */
        public short GetFootnoteInfo()
        {
            return field_3_footnoteInfo;
        }

        /**
         * Set the footnoteInfo field for the DOP record.
         */
        public void SetFootnoteInfo(short field_3_footnoteInfo)
        {
            this.field_3_footnoteInfo = field_3_footnoteInfo;
        }

        /**
         * Get the fOutlineDirtySave field for the DOP record.
         */
        public byte GetFOutlineDirtySave()
        {
            return field_4_fOutlineDirtySave;
        }

        /**
         * Set the fOutlineDirtySave field for the DOP record.
         */
        public void SetFOutlineDirtySave(byte field_4_fOutlineDirtySave)
        {
            this.field_4_fOutlineDirtySave = field_4_fOutlineDirtySave;
        }

        /**
         * Get the docinfo field for the DOP record.
         */
        public byte GetDocinfo()
        {
            return field_5_docinfo;
        }

        /**
         * Set the docinfo field for the DOP record.
         */
        public void SetDocinfo(byte field_5_docinfo)
        {
            this.field_5_docinfo = field_5_docinfo;
        }

        /**
         * Get the docinfo1 field for the DOP record.
         */
        public byte GetDocinfo1()
        {
            return field_6_docinfo1;
        }

        /**
         * Set the docinfo1 field for the DOP record.
         */
        public void SetDocinfo1(byte field_6_docinfo1)
        {
            this.field_6_docinfo1 = field_6_docinfo1;
        }

        /**
         * Get the docinfo2 field for the DOP record.
         */
        public byte GetDocinfo2()
        {
            return field_7_docinfo2;
        }

        /**
         * Set the docinfo2 field for the DOP record.
         */
        public void SetDocinfo2(byte field_7_docinfo2)
        {
            this.field_7_docinfo2 = field_7_docinfo2;
        }

        /**
         * Get the docinfo3 field for the DOP record.
         */
        public short GetDocinfo3()
        {
            return field_8_docinfo3;
        }

        /**
         * Set the docinfo3 field for the DOP record.
         */
        public void SetDocinfo3(short field_8_docinfo3)
        {
            this.field_8_docinfo3 = field_8_docinfo3;
        }

        /**
         * Get the dxaTab field for the DOP record.
         */
        public int GetDxaTab()
        {
            return field_9_dxaTab;
        }

        /**
         * Set the dxaTab field for the DOP record.
         */
        public void SetDxaTab(int field_9_dxaTab)
        {
            this.field_9_dxaTab = field_9_dxaTab;
        }

        /**
         * Get the wSpare field for the DOP record.
         */
        public int GetWSpare()
        {
            return field_10_wSpare;
        }

        /**
         * Set the wSpare field for the DOP record.
         */
        public void SetWSpare(int field_10_wSpare)
        {
            this.field_10_wSpare = field_10_wSpare;
        }

        /**
         * Get the dxaHotz field for the DOP record.
         */
        public int GetDxaHotz()
        {
            return field_11_dxaHotz;
        }

        /**
         * Set the dxaHotz field for the DOP record.
         */
        public void SetDxaHotz(int field_11_dxaHotz)
        {
            this.field_11_dxaHotz = field_11_dxaHotz;
        }

        /**
         * Get the cConsexHypLim field for the DOP record.
         */
        public int GetCConsexHypLim()
        {
            return field_12_cConsexHypLim;
        }

        /**
         * Set the cConsexHypLim field for the DOP record.
         */
        public void SetCConsexHypLim(int field_12_cConsexHypLim)
        {
            this.field_12_cConsexHypLim = field_12_cConsexHypLim;
        }

        /**
         * Get the wSpare2 field for the DOP record.
         */
        public int GetWSpare2()
        {
            return field_13_wSpare2;
        }

        /**
         * Set the wSpare2 field for the DOP record.
         */
        public void SetWSpare2(int field_13_wSpare2)
        {
            this.field_13_wSpare2 = field_13_wSpare2;
        }

        /**
         * Get the dttmCreated field for the DOP record.
         */
        public int GetDttmCreated()
        {
            return field_14_dttmCreated;
        }

        /**
         * Set the dttmCreated field for the DOP record.
         */
        public void SetDttmCreated(int field_14_dttmCreated)
        {
            this.field_14_dttmCreated = field_14_dttmCreated;
        }

        /**
         * Get the dttmRevised field for the DOP record.
         */
        public int GetDttmRevised()
        {
            return field_15_dttmRevised;
        }

        /**
         * Set the dttmRevised field for the DOP record.
         */
        public void SetDttmRevised(int field_15_dttmRevised)
        {
            this.field_15_dttmRevised = field_15_dttmRevised;
        }

        /**
         * Get the dttmLastPrint field for the DOP record.
         */
        public int GetDttmLastPrint()
        {
            return field_16_dttmLastPrint;
        }

        /**
         * Set the dttmLastPrint field for the DOP record.
         */
        public void SetDttmLastPrint(int field_16_dttmLastPrint)
        {
            this.field_16_dttmLastPrint = field_16_dttmLastPrint;
        }

        /**
         * Get the nRevision field for the DOP record.
         */
        public int GetNRevision()
        {
            return field_17_nRevision;
        }

        /**
         * Set the nRevision field for the DOP record.
         */
        public void SetNRevision(int field_17_nRevision)
        {
            this.field_17_nRevision = field_17_nRevision;
        }

        /**
         * Get the tmEdited field for the DOP record.
         */
        public int GetTmEdited()
        {
            return field_18_tmEdited;
        }

        /**
         * Set the tmEdited field for the DOP record.
         */
        public void SetTmEdited(int field_18_tmEdited)
        {
            this.field_18_tmEdited = field_18_tmEdited;
        }

        /**
         * Get the cWords field for the DOP record.
         */
        public int GetCWords()
        {
            return field_19_cWords;
        }

        /**
         * Set the cWords field for the DOP record.
         */
        public void SetCWords(int field_19_cWords)
        {
            this.field_19_cWords = field_19_cWords;
        }

        /**
         * Get the cCh field for the DOP record.
         */
        public int GetCCh()
        {
            return field_20_cCh;
        }

        /**
         * Set the cCh field for the DOP record.
         */
        public void SetCCh(int field_20_cCh)
        {
            this.field_20_cCh = field_20_cCh;
        }

        /**
         * Get the cPg field for the DOP record.
         */
        public int GetCPg()
        {
            return field_21_cPg;
        }

        /**
         * Set the cPg field for the DOP record.
         */
        public void SetCPg(int field_21_cPg)
        {
            this.field_21_cPg = field_21_cPg;
        }

        /**
         * Get the cParas field for the DOP record.
         */
        public int GetCParas()
        {
            return field_22_cParas;
        }

        /**
         * Set the cParas field for the DOP record.
         */
        public void SetCParas(int field_22_cParas)
        {
            this.field_22_cParas = field_22_cParas;
        }

        /**
         * Get the Edn field for the DOP record.
         */
        public short GetEdn()
        {
            return field_23_Edn;
        }

        /**
         * Set the Edn field for the DOP record.
         */
        public void SetEdn(short field_23_Edn)
        {
            this.field_23_Edn = field_23_Edn;
        }

        /**
         * Get the Edn1 field for the DOP record.
         */
        public short GetEdn1()
        {
            return field_24_Edn1;
        }

        /**
         * Set the Edn1 field for the DOP record.
         */
        public void SetEdn1(short field_24_Edn1)
        {
            this.field_24_Edn1 = field_24_Edn1;
        }

        /**
         * Get the cLines field for the DOP record.
         */
        public int GetCLines()
        {
            return field_25_cLines;
        }

        /**
         * Set the cLines field for the DOP record.
         */
        public void SetCLines(int field_25_cLines)
        {
            this.field_25_cLines = field_25_cLines;
        }

        /**
         * Get the cWordsFtnEnd field for the DOP record.
         */
        public int GetCWordsFtnEnd()
        {
            return field_26_cWordsFtnEnd;
        }

        /**
         * Set the cWordsFtnEnd field for the DOP record.
         */
        public void SetCWordsFtnEnd(int field_26_cWordsFtnEnd)
        {
            this.field_26_cWordsFtnEnd = field_26_cWordsFtnEnd;
        }

        /**
         * Get the cChFtnEdn field for the DOP record.
         */
        public int GetCChFtnEdn()
        {
            return field_27_cChFtnEdn;
        }

        /**
         * Set the cChFtnEdn field for the DOP record.
         */
        public void SetCChFtnEdn(int field_27_cChFtnEdn)
        {
            this.field_27_cChFtnEdn = field_27_cChFtnEdn;
        }

        /**
         * Get the cPgFtnEdn field for the DOP record.
         */
        public short GetCPgFtnEdn()
        {
            return field_28_cPgFtnEdn;
        }

        /**
         * Set the cPgFtnEdn field for the DOP record.
         */
        public void SetCPgFtnEdn(short field_28_cPgFtnEdn)
        {
            this.field_28_cPgFtnEdn = field_28_cPgFtnEdn;
        }

        /**
         * Get the cParasFtnEdn field for the DOP record.
         */
        public int GetCParasFtnEdn()
        {
            return field_29_cParasFtnEdn;
        }

        /**
         * Set the cParasFtnEdn field for the DOP record.
         */
        public void SetCParasFtnEdn(int field_29_cParasFtnEdn)
        {
            this.field_29_cParasFtnEdn = field_29_cParasFtnEdn;
        }

        /**
         * Get the cLinesFtnEdn field for the DOP record.
         */
        public int GetCLinesFtnEdn()
        {
            return field_30_cLinesFtnEdn;
        }

        /**
         * Set the cLinesFtnEdn field for the DOP record.
         */
        public void SetCLinesFtnEdn(int field_30_cLinesFtnEdn)
        {
            this.field_30_cLinesFtnEdn = field_30_cLinesFtnEdn;
        }

        /**
         * Get the lKeyProtDoc field for the DOP record.
         */
        public int GetLKeyProtDoc()
        {
            return field_31_lKeyProtDoc;
        }

        /**
         * Set the lKeyProtDoc field for the DOP record.
         */
        public void SetLKeyProtDoc(int field_31_lKeyProtDoc)
        {
            this.field_31_lKeyProtDoc = field_31_lKeyProtDoc;
        }

        /**
         * Get the view field for the DOP record.
         */
        public short GetView()
        {
            return field_32_view;
        }

        /**
         * Set the view field for the DOP record.
         */
        public void SetView(short field_32_view)
        {
            this.field_32_view = field_32_view;
        }

        /**
         * Get the docinfo4 field for the DOP record.
         */
        public int GetDocinfo4()
        {
            return field_33_docinfo4;
        }

        /**
         * Set the docinfo4 field for the DOP record.
         */
        public void SetDocinfo4(int field_33_docinfo4)
        {
            this.field_33_docinfo4 = field_33_docinfo4;
        }

        /**
         * Get the adt field for the DOP record.
         */
        public short GetAdt()
        {
            return field_34_adt;
        }

        /**
         * Set the adt field for the DOP record.
         */
        public void SetAdt(short field_34_adt)
        {
            this.field_34_adt = field_34_adt;
        }

        /**
         * Get the doptypography field for the DOP record.
         */
        public byte[] GetDoptypography()
        {
            return field_35_doptypography;
        }

        /**
         * Set the doptypography field for the DOP record.
         */
        public void SetDoptypography(byte[] field_35_doptypography)
        {
            this.field_35_doptypography = field_35_doptypography;
        }

        /**
         * Get the dogrid field for the DOP record.
         */
        public byte[] GetDogrid()
        {
            return field_36_dogrid;
        }

        /**
         * Set the dogrid field for the DOP record.
         */
        public void SetDogrid(byte[] field_36_dogrid)
        {
            this.field_36_dogrid = field_36_dogrid;
        }

        /**
         * Get the docinfo5 field for the DOP record.
         */
        public short GetDocinfo5()
        {
            return field_37_docinfo5;
        }

        /**
         * Set the docinfo5 field for the DOP record.
         */
        public void SetDocinfo5(short field_37_docinfo5)
        {
            this.field_37_docinfo5 = field_37_docinfo5;
        }

        /**
         * Get the docinfo6 field for the DOP record.
         */
        public short GetDocinfo6()
        {
            return field_38_docinfo6;
        }

        /**
         * Set the docinfo6 field for the DOP record.
         */
        public void SetDocinfo6(short field_38_docinfo6)
        {
            this.field_38_docinfo6 = field_38_docinfo6;
        }

        /**
         * Get the asumyi field for the DOP record.
         */
        public byte[] GetAsumyi()
        {
            return field_39_asumyi;
        }

        /**
         * Set the asumyi field for the DOP record.
         */
        public void SetAsumyi(byte[] field_39_asumyi)
        {
            this.field_39_asumyi = field_39_asumyi;
        }

        /**
         * Get the cChWS field for the DOP record.
         */
        public int GetCChWS()
        {
            return field_40_cChWS;
        }

        /**
         * Set the cChWS field for the DOP record.
         */
        public void SetCChWS(int field_40_cChWS)
        {
            this.field_40_cChWS = field_40_cChWS;
        }

        /**
         * Get the cChWSFtnEdn field for the DOP record.
         */
        public int GetCChWSFtnEdn()
        {
            return field_41_cChWSFtnEdn;
        }

        /**
         * Set the cChWSFtnEdn field for the DOP record.
         */
        public void SetCChWSFtnEdn(int field_41_cChWSFtnEdn)
        {
            this.field_41_cChWSFtnEdn = field_41_cChWSFtnEdn;
        }

        /**
         * Get the grfDocEvents field for the DOP record.
         */
        public int GetGrfDocEvents()
        {
            return field_42_grfDocEvents;
        }

        /**
         * Set the grfDocEvents field for the DOP record.
         */
        public void SetGrfDocEvents(int field_42_grfDocEvents)
        {
            this.field_42_grfDocEvents = field_42_grfDocEvents;
        }

        /**
         * Get the virusinfo field for the DOP record.
         */
        public int GetVirusinfo()
        {
            return field_43_virusinfo;
        }

        /**
         * Set the virusinfo field for the DOP record.
         */
        public void SetVirusinfo(int field_43_virusinfo)
        {
            this.field_43_virusinfo = field_43_virusinfo;
        }

        /**
         * Get the Spare field for the DOP record.
         */
        public byte[] GetSpare()
        {
            return field_44_Spare;
        }

        /**
         * Set the Spare field for the DOP record.
         */
        public void SetSpare(byte[] field_44_Spare)
        {
            this.field_44_Spare = field_44_Spare;
        }

        /**
         * Get the reserved1 field for the DOP record.
         */
        public int GetReserved1()
        {
            return field_45_reserved1;
        }

        /**
         * Set the reserved1 field for the DOP record.
         */
        public void SetReserved1(int field_45_reserved1)
        {
            this.field_45_reserved1 = field_45_reserved1;
        }

        /**
         * Get the reserved2 field for the DOP record.
         */
        public int GetReserved2()
        {
            return field_46_reserved2;
        }

        /**
         * Set the reserved2 field for the DOP record.
         */
        public void SetReserved2(int field_46_reserved2)
        {
            this.field_46_reserved2 = field_46_reserved2;
        }

        /**
         * Get the cDBC field for the DOP record.
         */
        public int GetCDBC()
        {
            return field_47_cDBC;
        }

        /**
         * Set the cDBC field for the DOP record.
         */
        public void SetCDBC(int field_47_cDBC)
        {
            this.field_47_cDBC = field_47_cDBC;
        }

        /**
         * Get the cDBCFtnEdn field for the DOP record.
         */
        public int GetCDBCFtnEdn()
        {
            return field_48_cDBCFtnEdn;
        }

        /**
         * Set the cDBCFtnEdn field for the DOP record.
         */
        public void SetCDBCFtnEdn(int field_48_cDBCFtnEdn)
        {
            this.field_48_cDBCFtnEdn = field_48_cDBCFtnEdn;
        }

        /**
         * Get the reserved field for the DOP record.
         */
        public int GetReserved()
        {
            return field_49_reserved;
        }

        /**
         * Set the reserved field for the DOP record.
         */
        public void SetReserved(int field_49_reserved)
        {
            this.field_49_reserved = field_49_reserved;
        }

        /**
         * Get the nfcFtnRef field for the DOP record.
         */
        public short GetNfcFtnRef()
        {
            return field_50_nfcFtnRef;
        }

        /**
         * Set the nfcFtnRef field for the DOP record.
         */
        public void SetNfcFtnRef(short field_50_nfcFtnRef)
        {
            this.field_50_nfcFtnRef = field_50_nfcFtnRef;
        }

        /**
         * Get the nfcEdnRef field for the DOP record.
         */
        public short GetNfcEdnRef()
        {
            return field_51_nfcEdnRef;
        }

        /**
         * Set the nfcEdnRef field for the DOP record.
         */
        public void SetNfcEdnRef(short field_51_nfcEdnRef)
        {
            this.field_51_nfcEdnRef = field_51_nfcEdnRef;
        }

        /**
         * Get the hpsZoonFontPag field for the DOP record.
         */
        public short GetHpsZoonFontPag()
        {
            return field_52_hpsZoonFontPag;
        }

        /**
         * Set the hpsZoonFontPag field for the DOP record.
         */
        public void SetHpsZoonFontPag(short field_52_hpsZoonFontPag)
        {
            this.field_52_hpsZoonFontPag = field_52_hpsZoonFontPag;
        }

        /**
         * Get the dywDispPag field for the DOP record.
         */
        public short GetDywDispPag()
        {
            return field_53_dywDispPag;
        }

        /**
         * Set the dywDispPag field for the DOP record.
         */
        public void SetDywDispPag(short field_53_dywDispPag)
        {
            this.field_53_dywDispPag = field_53_dywDispPag;
        }

        /**
         * Sets the fFacingPages field value.
         * 
         */
        public void SetFFacingPages(bool value)
        {
            field_1_formatFlags = (byte)fFacingPages.SetBoolean(field_1_formatFlags, value);


        }

        /**
         * 
         * @return  the fFacingPages field value.
         */
        public bool IsFFacingPages()
        {
            return fFacingPages.IsSet(field_1_formatFlags);

        }

        /**
         * Sets the fWidowControl field value.
         * 
         */
        public void SetFWidowControl(bool value)
        {
            field_1_formatFlags = (byte)fWidowControl.SetBoolean(field_1_formatFlags, value);


        }

        /**
         * 
         * @return  the fWidowControl field value.
         */
        public bool IsFWidowControl()
        {
            return fWidowControl.IsSet(field_1_formatFlags);

        }

        /**
         * Sets the fPMHMainDoc field value.
         * 
         */
        public void SetFPMHMainDoc(bool value)
        {
            field_1_formatFlags = (byte)fPMHMainDoc.SetBoolean(field_1_formatFlags, value);


        }

        /**
         * 
         * @return  the fPMHMainDoc field value.
         */
        public bool IsFPMHMainDoc()
        {
            return fPMHMainDoc.IsSet(field_1_formatFlags);

        }

        /**
         * Sets the grfSupression field value.
         * 
         */
        public void SetGrfSupression(byte value)
        {
            field_1_formatFlags = (byte)grfSupression.SetValue(field_1_formatFlags, value);


        }

        /**
         * 
         * @return  the grfSupression field value.
         */
        public byte GetGrfSupression()
        {
            return (byte)grfSupression.GetValue(field_1_formatFlags);

        }

        /**
         * Sets the fpc field value.
         * 
         */
        public void SetFpc(byte value)
        {
            field_1_formatFlags = (byte)fpc.SetValue(field_1_formatFlags, value);


        }

        /**
         * 
         * @return  the fpc field value.
         */
        public byte GetFpc()
        {
            return (byte)fpc.GetValue(field_1_formatFlags);

        }

        /**
         * Sets the unused1 field value.
         * 
         */
        public void SetUnused1(bool value)
        {
            field_1_formatFlags = (byte)unused1.SetBoolean(field_1_formatFlags, value);


        }

        /**
         * 
         * @return  the unused1 field value.
         */
        public bool IsUnused1()
        {
            return unused1.IsSet(field_1_formatFlags);

        }

        /**
         * Sets the rncFtn field value.
         * 
         */
        public void SetRncFtn(byte value)
        {
            field_3_footnoteInfo = (short)rncFtn.SetValue(field_3_footnoteInfo, value);


        }

        /**
         * 
         * @return  the rncFtn field value.
         */
        public byte GetRncFtn()
        {
            return (byte)rncFtn.GetValue(field_3_footnoteInfo);

        }

        /**
         * Sets the nFtn field value.
         * 
         */
        public void SetNFtn(short value)
        {
            field_3_footnoteInfo = (short)nFtn.SetValue(field_3_footnoteInfo, value);


        }

        /**
         * 
         * @return  the nFtn field value.
         */
        public short GetNFtn()
        {
            return (short)nFtn.GetValue(field_3_footnoteInfo);

        }

        /**
         * Sets the fOnlyMacPics field value.
         * 
         */
        public void SetFOnlyMacPics(bool value)
        {
            field_5_docinfo = (byte)fOnlyMacPics.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fOnlyMacPics field value.
         */
        public bool IsFOnlyMacPics()
        {
            return fOnlyMacPics.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fOnlyWinPics field value.
         * 
         */
        public void SetFOnlyWinPics(bool value)
        {
            field_5_docinfo = (byte)fOnlyWinPics.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fOnlyWinPics field value.
         */
        public bool IsFOnlyWinPics()
        {
            return fOnlyWinPics.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fLabelDoc field value.
         * 
         */
        public void SetFLabelDoc(bool value)
        {
            field_5_docinfo = (byte)fLabelDoc.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fLabelDoc field value.
         */
        public bool IsFLabelDoc()
        {
            return fLabelDoc.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fHyphCapitals field value.
         * 
         */
        public void SetFHyphCapitals(bool value)
        {
            field_5_docinfo = (byte)fHyphCapitals.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fHyphCapitals field value.
         */
        public bool IsFHyphCapitals()
        {
            return fHyphCapitals.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fAutoHyphen field value.
         * 
         */
        public void SetFAutoHyphen(bool value)
        {
            field_5_docinfo = (byte)fAutoHyphen.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fAutoHyphen field value.
         */
        public bool IsFAutoHyphen()
        {
            return fAutoHyphen.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fFormNoFields field value.
         * 
         */
        public void SetFFormNoFields(bool value)
        {
            field_5_docinfo = (byte)fFormNoFields.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fFormNoFields field value.
         */
        public bool IsFFormNoFields()
        {
            return fFormNoFields.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fLinkStyles field value.
         * 
         */
        public void SetFLinkStyles(bool value)
        {
            field_5_docinfo = (byte)fLinkStyles.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fLinkStyles field value.
         */
        public bool IsFLinkStyles()
        {
            return fLinkStyles.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fRevMarking field value.
         * 
         */
        public void SetFRevMarking(bool value)
        {
            field_5_docinfo = (byte)fRevMarking.SetBoolean(field_5_docinfo, value);


        }

        /**
         * 
         * @return  the fRevMarking field value.
         */
        public bool IsFRevMarking()
        {
            return fRevMarking.IsSet(field_5_docinfo);

        }

        /**
         * Sets the fBackup field value.
         * 
         */
        public void SetFBackup(bool value)
        {
            field_6_docinfo1 = (byte)fBackup.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the fBackup field value.
         */
        public bool IsFBackup()
        {
            return fBackup.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the fExactCWords field value.
         * 
         */
        public void SetFExactCWords(bool value)
        {
            field_6_docinfo1 = (byte)fExactCWords.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the fExactCWords field value.
         */
        public bool IsFExactCWords()
        {
            return fExactCWords.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the fPagHidden field value.
         * 
         */
        public void SetFPagHidden(bool value)
        {
            field_6_docinfo1 = (byte)fPagHidden.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the fPagHidden field value.
         */
        public bool IsFPagHidden()
        {
            return fPagHidden.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the fPagResults field value.
         * 
         */
        public void SetFPagResults(bool value)
        {
            field_6_docinfo1 = (byte)fPagResults.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the fPagResults field value.
         */
        public bool IsFPagResults()
        {
            return fPagResults.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the fLockAtn field value.
         * 
         */
        public void SetFLockAtn(bool value)
        {
            field_6_docinfo1 = (byte)fLockAtn.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the fLockAtn field value.
         */
        public bool IsFLockAtn()
        {
            return fLockAtn.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the fMirrorMargins field value.
         * 
         */
        public void SetFMirrorMargins(bool value)
        {
            field_6_docinfo1 = (byte)fMirrorMargins.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the fMirrorMargins field value.
         */
        public bool IsFMirrorMargins()
        {
            return fMirrorMargins.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the unused3 field value.
         * 
         */
        public void SetUnused3(bool value)
        {
            field_6_docinfo1 = (byte)unused3.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the unused3 field value.
         */
        public bool IsUnused3()
        {
            return unused3.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the fDfltTrueType field value.
         * 
         */
        public void SetFDfltTrueType(bool value)
        {
            field_6_docinfo1 = (byte)fDfltTrueType.SetBoolean(field_6_docinfo1, value);


        }

        /**
         * 
         * @return  the fDfltTrueType field value.
         */
        public bool IsFDfltTrueType()
        {
            return fDfltTrueType.IsSet(field_6_docinfo1);

        }

        /**
         * Sets the fPagSupressTopSpacing field value.
         * 
         */
        public void SetFPagSupressTopSpacing(bool value)
        {
            field_7_docinfo2 = (byte)fPagSupressTopSpacing.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the fPagSupressTopSpacing field value.
         */
        public bool IsFPagSupressTopSpacing()
        {
            return fPagSupressTopSpacing.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the fProtEnabled field value.
         * 
         */
        public void SetFProtEnabled(bool value)
        {
            field_7_docinfo2 = (byte)fProtEnabled.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the fProtEnabled field value.
         */
        public bool IsFProtEnabled()
        {
            return fProtEnabled.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the fDispFormFldSel field value.
         * 
         */
        public void SetFDispFormFldSel(bool value)
        {
            field_7_docinfo2 = (byte)fDispFormFldSel.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the fDispFormFldSel field value.
         */
        public bool IsFDispFormFldSel()
        {
            return fDispFormFldSel.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the fRMView field value.
         * 
         */
        public void SetFRMView(bool value)
        {
            field_7_docinfo2 = (byte)fRMView.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the fRMView field value.
         */
        public bool IsFRMView()
        {
            return fRMView.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the fRMPrint field value.
         * 
         */
        public void SetFRMPrint(bool value)
        {
            field_7_docinfo2 = (byte)fRMPrint.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the fRMPrint field value.
         */
        public bool IsFRMPrint()
        {
            return fRMPrint.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the unused4 field value.
         * 
         */
        public void SetUnused4(bool value)
        {
            field_7_docinfo2 = (byte)unused4.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the unused4 field value.
         */
        public bool IsUnused4()
        {
            return unused4.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the fLockRev field value.
         * 
         */
        public void SetFLockRev(bool value)
        {
            field_7_docinfo2 = (byte)fLockRev.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the fLockRev field value.
         */
        public bool IsFLockRev()
        {
            return fLockRev.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the fEmbedFonts field value.
         * 
         */
        public void SetFEmbedFonts(bool value)
        {
            field_7_docinfo2 = (byte)fEmbedFonts.SetBoolean(field_7_docinfo2, value);


        }

        /**
         * 
         * @return  the fEmbedFonts field value.
         */
        public bool IsFEmbedFonts()
        {
            return fEmbedFonts.IsSet(field_7_docinfo2);

        }

        /**
         * Sets the oldfNoTabForInd field value.
         * 
         */
        public void SetOldfNoTabForInd(bool value)
        {
            field_8_docinfo3 = (short)oldfNoTabForInd.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfNoTabForInd field value.
         */
        public bool IsOldfNoTabForInd()
        {
            return oldfNoTabForInd.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfNoSpaceRaiseLower field value.
         * 
         */
        public void SetOldfNoSpaceRaiseLower(bool value)
        {
            field_8_docinfo3 = (short)oldfNoSpaceRaiseLower.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfNoSpaceRaiseLower field value.
         */
        public bool IsOldfNoSpaceRaiseLower()
        {
            return oldfNoSpaceRaiseLower.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfSuppressSpbfAfterPageBreak field value.
         * 
         */
        public void SetOldfSuppressSpbfAfterPageBreak(bool value)
        {
            field_8_docinfo3 = (short)oldfSuppressSpbfAfterPageBreak.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfSuppressSpbfAfterPageBreak field value.
         */
        public bool IsOldfSuppressSpbfAfterPageBreak()
        {
            return oldfSuppressSpbfAfterPageBreak.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfWrapTrailSpaces field value.
         * 
         */
        public void SetOldfWrapTrailSpaces(bool value)
        {
            field_8_docinfo3 = (short)oldfWrapTrailSpaces.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfWrapTrailSpaces field value.
         */
        public bool IsOldfWrapTrailSpaces()
        {
            return oldfWrapTrailSpaces.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfMapPrintTextColor field value.
         * 
         */
        public void SetOldfMapPrintTextColor(bool value)
        {
            field_8_docinfo3 = (short)oldfMapPrintTextColor.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfMapPrintTextColor field value.
         */
        public bool IsOldfMapPrintTextColor()
        {
            return oldfMapPrintTextColor.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfNoColumnBalance field value.
         * 
         */
        public void SetOldfNoColumnBalance(bool value)
        {
            field_8_docinfo3 = (short)oldfNoColumnBalance.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfNoColumnBalance field value.
         */
        public bool IsOldfNoColumnBalance()
        {
            return oldfNoColumnBalance.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfConvMailMergeEsc field value.
         * 
         */
        public void SetOldfConvMailMergeEsc(bool value)
        {
            field_8_docinfo3 = (short)oldfConvMailMergeEsc.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfConvMailMergeEsc field value.
         */
        public bool IsOldfConvMailMergeEsc()
        {
            return oldfConvMailMergeEsc.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfSupressTopSpacing field value.
         * 
         */
        public void SetOldfSupressTopSpacing(bool value)
        {
            field_8_docinfo3 = (short)oldfSupressTopSpacing.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfSupressTopSpacing field value.
         */
        public bool IsOldfSupressTopSpacing()
        {
            return oldfSupressTopSpacing.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfOrigWordTableRules field value.
         * 
         */
        public void SetOldfOrigWordTableRules(bool value)
        {
            field_8_docinfo3 = (short)oldfOrigWordTableRules.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfOrigWordTableRules field value.
         */
        public bool IsOldfOrigWordTableRules()
        {
            return oldfOrigWordTableRules.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfTransparentMetafiles field value.
         * 
         */
        public void SetOldfTransparentMetafiles(bool value)
        {
            field_8_docinfo3 = (short)oldfTransparentMetafiles.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfTransparentMetafiles field value.
         */
        public bool IsOldfTransparentMetafiles()
        {
            return oldfTransparentMetafiles.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfShowBreaksInFrames field value.
         * 
         */
        public void SetOldfShowBreaksInFrames(bool value)
        {
            field_8_docinfo3 = (short)oldfShowBreaksInFrames.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfShowBreaksInFrames field value.
         */
        public bool IsOldfShowBreaksInFrames()
        {
            return oldfShowBreaksInFrames.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the oldfSwapBordersFacingPgs field value.
         * 
         */
        public void SetOldfSwapBordersFacingPgs(bool value)
        {
            field_8_docinfo3 = (short)oldfSwapBordersFacingPgs.SetBoolean(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the oldfSwapBordersFacingPgs field value.
         */
        public bool IsOldfSwapBordersFacingPgs()
        {
            return oldfSwapBordersFacingPgs.IsSet(field_8_docinfo3);

        }

        /**
         * Sets the unused5 field value.
         * 
         */
        public void SetUnused5(byte value)
        {
            field_8_docinfo3 = (short)unused5.SetValue(field_8_docinfo3, value);


        }

        /**
         * 
         * @return  the unused5 field value.
         */
        public byte GetUnused5()
        {
            return (byte)unused5.GetValue(field_8_docinfo3);

        }

        /**
         * Sets the rncEdn field value.
         * 
         */
        public void SetRncEdn(byte value)
        {
            field_23_Edn = (short)rncEdn.SetValue(field_23_Edn, value);


        }

        /**
         * 
         * @return  the rncEdn field value.
         */
        public byte GetRncEdn()
        {
            return (byte)rncEdn.GetValue(field_23_Edn);

        }

        /**
         * Sets the nEdn field value.
         * 
         */
        public void SetNEdn(short value)
        {
            field_23_Edn = (short)nEdn.SetValue(field_23_Edn, value);


        }

        /**
         * 
         * @return  the nEdn field value.
         */
        public short GetNEdn()
        {
            return (short)nEdn.GetValue(field_23_Edn);

        }

        /**
         * Sets the epc field value.
         * 
         */
        public void SetEpc(byte value)
        {
            field_24_Edn1 = (short)epc.SetValue(field_24_Edn1, value);


        }

        /**
         * 
         * @return  the epc field value.
         */
        public byte GetEpc()
        {
            return (byte)epc.GetValue(field_24_Edn1);

        }

        /**
         * Sets the nfcFtnRef1 field value.
         * 
         */
        public void SetNfcFtnRef1(byte value)
        {
            field_24_Edn1 = (short)nfcFtnRef1.SetValue(field_24_Edn1, value);


        }

        /**
         * 
         * @return  the nfcFtnRef1 field value.
         */
        public byte GetNfcFtnRef1()
        {
            return (byte)nfcFtnRef1.GetValue(field_24_Edn1);

        }

        /**
         * Sets the nfcEdnRef1 field value.
         * 
         */
        public void SetNfcEdnRef1(byte value)
        {
            field_24_Edn1 = (short)nfcEdnRef1.SetValue(field_24_Edn1, value);


        }

        /**
         * 
         * @return  the nfcEdnRef1 field value.
         */
        public byte GetNfcEdnRef1()
        {
            return (byte)nfcEdnRef1.GetValue(field_24_Edn1);

        }

        /**
         * Sets the fPrintFormData field value.
         * 
         */
        public void SetFPrintFormData(bool value)
        {
            field_24_Edn1 = (short)fPrintFormData.SetBoolean(field_24_Edn1, value);


        }

        /**
         * 
         * @return  the fPrintFormData field value.
         */
        public bool IsFPrintFormData()
        {
            return fPrintFormData.IsSet(field_24_Edn1);

        }

        /**
         * Sets the fSaveFormData field value.
         * 
         */
        public void SetFSaveFormData(bool value)
        {
            field_24_Edn1 = (short)fSaveFormData.SetBoolean(field_24_Edn1, value);


        }

        /**
         * 
         * @return  the fSaveFormData field value.
         */
        public bool IsFSaveFormData()
        {
            return fSaveFormData.IsSet(field_24_Edn1);

        }

        /**
         * Sets the fShadeFormData field value.
         * 
         */
        public void SetFShadeFormData(bool value)
        {
            field_24_Edn1 = (short)fShadeFormData.SetBoolean(field_24_Edn1, value);


        }

        /**
         * 
         * @return  the fShadeFormData field value.
         */
        public bool IsFShadeFormData()
        {
            return fShadeFormData.IsSet(field_24_Edn1);

        }

        /**
         * Sets the fWCFtnEdn field value.
         * 
         */
        public void SetFWCFtnEdn(bool value)
        {
            field_24_Edn1 = (short)fWCFtnEdn.SetBoolean(field_24_Edn1, value);


        }

        /**
         * 
         * @return  the fWCFtnEdn field value.
         */
        public bool IsFWCFtnEdn()
        {
            return fWCFtnEdn.IsSet(field_24_Edn1);

        }

        /**
         * Sets the wvkSaved field value.
         * 
         */
        public void SetWvkSaved(byte value)
        {
            field_32_view = (short)wvkSaved.SetValue(field_32_view, value);


        }

        /**
         * 
         * @return  the wvkSaved field value.
         */
        public byte GetWvkSaved()
        {
            return (byte)wvkSaved.GetValue(field_32_view);

        }

        /**
         * Sets the wScaleSaved field value.
         * 
         */
        public void SetWScaleSaved(short value)
        {
            field_32_view = (short)wScaleSaved.SetValue(field_32_view, value);


        }

        /**
         * 
         * @return  the wScaleSaved field value.
         */
        public short GetWScaleSaved()
        {
            return (short)wScaleSaved.GetValue(field_32_view);

        }

        /**
         * Sets the zkSaved field value.
         * 
         */
        public void SetZkSaved(byte value)
        {
            field_32_view = (short)zkSaved.SetValue(field_32_view, value);


        }

        /**
         * 
         * @return  the zkSaved field value.
         */
        public byte GetZkSaved()
        {
            return (byte)zkSaved.GetValue(field_32_view);

        }

        /**
         * Sets the fRotateFontW6 field value.
         * 
         */
        public void SetFRotateFontW6(bool value)
        {
            field_32_view = (short)fRotateFontW6.SetBoolean(field_32_view, value);


        }

        /**
         * 
         * @return  the fRotateFontW6 field value.
         */
        public bool IsFRotateFontW6()
        {
            return fRotateFontW6.IsSet(field_32_view);

        }

        /**
         * Sets the iGutterPos field value.
         * 
         */
        public void SetIGutterPos(bool value)
        {
            field_32_view = (short)iGutterPos.SetBoolean(field_32_view, value);


        }

        /**
         * 
         * @return  the iGutterPos field value.
         */
        public bool IsIGutterPos()
        {
            return iGutterPos.IsSet(field_32_view);

        }

        /**
         * Sets the fNoTabForInd field value.
         * 
         */
        public void SetFNoTabForInd(bool value)
        {
            field_33_docinfo4 = (int)fNoTabForInd.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fNoTabForInd field value.
         */
        public bool IsFNoTabForInd()
        {
            return fNoTabForInd.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fNoSpaceRaiseLower field value.
         * 
         */
        public void SetFNoSpaceRaiseLower(bool value)
        {
            field_33_docinfo4 = (int)fNoSpaceRaiseLower.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fNoSpaceRaiseLower field value.
         */
        public bool IsFNoSpaceRaiseLower()
        {
            return fNoSpaceRaiseLower.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fSupressSpdfAfterPageBreak field value.
         * 
         */
        public void SetFSupressSpdfAfterPageBreak(bool value)
        {
            field_33_docinfo4 = (int)fSupressSpdfAfterPageBreak.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fSupressSpdfAfterPageBreak field value.
         */
        public bool IsFSupressSpdfAfterPageBreak()
        {
            return fSupressSpdfAfterPageBreak.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fWrapTrailSpaces field value.
         * 
         */
        public void SetFWrapTrailSpaces(bool value)
        {
            field_33_docinfo4 = (int)fWrapTrailSpaces.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fWrapTrailSpaces field value.
         */
        public bool IsFWrapTrailSpaces()
        {
            return fWrapTrailSpaces.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fMapPrintTextColor field value.
         * 
         */
        public void SetFMapPrintTextColor(bool value)
        {
            field_33_docinfo4 = (int)fMapPrintTextColor.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fMapPrintTextColor field value.
         */
        public bool IsFMapPrintTextColor()
        {
            return fMapPrintTextColor.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fNoColumnBalance field value.
         * 
         */
        public void SetFNoColumnBalance(bool value)
        {
            field_33_docinfo4 = (int)fNoColumnBalance.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fNoColumnBalance field value.
         */
        public bool IsFNoColumnBalance()
        {
            return fNoColumnBalance.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fConvMailMergeEsc field value.
         * 
         */
        public void SetFConvMailMergeEsc(bool value)
        {
            field_33_docinfo4 = (int)fConvMailMergeEsc.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fConvMailMergeEsc field value.
         */
        public bool IsFConvMailMergeEsc()
        {
            return fConvMailMergeEsc.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fSupressTopSpacing field value.
         * 
         */
        public void SetFSupressTopSpacing(bool value)
        {
            field_33_docinfo4 = (int)fSupressTopSpacing.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fSupressTopSpacing field value.
         */
        public bool IsFSupressTopSpacing()
        {
            return fSupressTopSpacing.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fOrigWordTableRules field value.
         * 
         */
        public void SetFOrigWordTableRules(bool value)
        {
            field_33_docinfo4 = (int)fOrigWordTableRules.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fOrigWordTableRules field value.
         */
        public bool IsFOrigWordTableRules()
        {
            return fOrigWordTableRules.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fTransparentMetafiles field value.
         * 
         */
        public void SetFTransparentMetafiles(bool value)
        {
            field_33_docinfo4 = (int)fTransparentMetafiles.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fTransparentMetafiles field value.
         */
        public bool IsFTransparentMetafiles()
        {
            return fTransparentMetafiles.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fShowBreaksInFrames field value.
         * 
         */
        public void SetFShowBreaksInFrames(bool value)
        {
            field_33_docinfo4 = (int)fShowBreaksInFrames.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fShowBreaksInFrames field value.
         */
        public bool IsFShowBreaksInFrames()
        {
            return fShowBreaksInFrames.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fSwapBordersFacingPgs field value.
         * 
         */
        public void SetFSwapBordersFacingPgs(bool value)
        {
            field_33_docinfo4 = (int)fSwapBordersFacingPgs.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fSwapBordersFacingPgs field value.
         */
        public bool IsFSwapBordersFacingPgs()
        {
            return fSwapBordersFacingPgs.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fSuppressTopSPacingMac5 field value.
         * 
         */
        public void SetFSuppressTopSPacingMac5(bool value)
        {
            field_33_docinfo4 = (int)fSuppressTopSPacingMac5.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fSuppressTopSPacingMac5 field value.
         */
        public bool IsFSuppressTopSPacingMac5()
        {
            return fSuppressTopSPacingMac5.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fTruncDxaExpand field value.
         * 
         */
        public void SetFTruncDxaExpand(bool value)
        {
            field_33_docinfo4 = (int)fTruncDxaExpand.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fTruncDxaExpand field value.
         */
        public bool IsFTruncDxaExpand()
        {
            return fTruncDxaExpand.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fPrintBodyBeforeHdr field value.
         * 
         */
        public void SetFPrintBodyBeforeHdr(bool value)
        {
            field_33_docinfo4 = (int)fPrintBodyBeforeHdr.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fPrintBodyBeforeHdr field value.
         */
        public bool IsFPrintBodyBeforeHdr()
        {
            return fPrintBodyBeforeHdr.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fNoLeading field value.
         * 
         */
        public void SetFNoLeading(bool value)
        {
            field_33_docinfo4 = (int)fNoLeading.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fNoLeading field value.
         */
        public bool IsFNoLeading()
        {
            return fNoLeading.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the fMWSmallCaps field value.
         * 
         */
        public void SetFMWSmallCaps(bool value)
        {
            field_33_docinfo4 = (int)fMWSmallCaps.SetBoolean(field_33_docinfo4, value);


        }

        /**
         * 
         * @return  the fMWSmallCaps field value.
         */
        public bool IsFMWSmallCaps()
        {
            return fMWSmallCaps.IsSet(field_33_docinfo4);

        }

        /**
         * Sets the lvl field value.
         * 
         */
        public void SetLvl(byte value)
        {
            field_37_docinfo5 = (short)lvl.SetValue(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the lvl field value.
         */
        public byte GetLvl()
        {
            return (byte)lvl.GetValue(field_37_docinfo5);

        }

        /**
         * Sets the fGramAllDone field value.
         * 
         */
        public void SetFGramAllDone(bool value)
        {
            field_37_docinfo5 = (short)fGramAllDone.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fGramAllDone field value.
         */
        public bool IsFGramAllDone()
        {
            return fGramAllDone.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fGramAllClean field value.
         * 
         */
        public void SetFGramAllClean(bool value)
        {
            field_37_docinfo5 = (short)fGramAllClean.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fGramAllClean field value.
         */
        public bool IsFGramAllClean()
        {
            return fGramAllClean.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fSubsetFonts field value.
         * 
         */
        public void SetFSubsetFonts(bool value)
        {
            field_37_docinfo5 = (short)fSubsetFonts.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fSubsetFonts field value.
         */
        public bool IsFSubsetFonts()
        {
            return fSubsetFonts.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fHideLastVersion field value.
         * 
         */
        public void SetFHideLastVersion(bool value)
        {
            field_37_docinfo5 = (short)fHideLastVersion.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fHideLastVersion field value.
         */
        public bool IsFHideLastVersion()
        {
            return fHideLastVersion.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fHtmlDoc field value.
         * 
         */
        public void SetFHtmlDoc(bool value)
        {
            field_37_docinfo5 = (short)fHtmlDoc.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fHtmlDoc field value.
         */
        public bool IsFHtmlDoc()
        {
            return fHtmlDoc.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fSnapBorder field value.
         * 
         */
        public void SetFSnapBorder(bool value)
        {
            field_37_docinfo5 = (short)fSnapBorder.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fSnapBorder field value.
         */
        public bool IsFSnapBorder()
        {
            return fSnapBorder.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fIncludeHeader field value.
         * 
         */
        public void SetFIncludeHeader(bool value)
        {
            field_37_docinfo5 = (short)fIncludeHeader.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fIncludeHeader field value.
         */
        public bool IsFIncludeHeader()
        {
            return fIncludeHeader.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fIncludeFooter field value.
         * 
         */
        public void SetFIncludeFooter(bool value)
        {
            field_37_docinfo5 = (short)fIncludeFooter.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fIncludeFooter field value.
         */
        public bool IsFIncludeFooter()
        {
            return fIncludeFooter.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fForcePageSizePag field value.
         * 
         */
        public void SetFForcePageSizePag(bool value)
        {
            field_37_docinfo5 = (short)fForcePageSizePag.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fForcePageSizePag field value.
         */
        public bool IsFForcePageSizePag()
        {
            return fForcePageSizePag.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fMinFontSizePag field value.
         * 
         */
        public void SetFMinFontSizePag(bool value)
        {
            field_37_docinfo5 = (short)fMinFontSizePag.SetBoolean(field_37_docinfo5, value);


        }

        /**
         * 
         * @return  the fMinFontSizePag field value.
         */
        public bool IsFMinFontSizePag()
        {
            return fMinFontSizePag.IsSet(field_37_docinfo5);

        }

        /**
         * Sets the fHaveVersions field value.
         * 
         */
        public void SetFHaveVersions(bool value)
        {
            field_38_docinfo6 = (short)fHaveVersions.SetBoolean(field_38_docinfo6, value);


        }

        /**
         * 
         * @return  the fHaveVersions field value.
         */
        public bool IsFHaveVersions()
        {
            return fHaveVersions.IsSet(field_38_docinfo6);

        }

        /**
         * Sets the fAutoVersions field value.
         * 
         */
        public void SetFAutoVersions(bool value)
        {
            field_38_docinfo6 = (short)fAutoVersions.SetBoolean(field_38_docinfo6, value);


        }

        /**
         * 
         * @return  the fAutoVersions field value.
         */
        public bool IsFAutoVersions()
        {
            return fAutoVersions.IsSet(field_38_docinfo6);

        }

        /**
         * Sets the fVirusPrompted field value.
         * 
         */
        public void SetFVirusPrompted(bool value)
        {
            field_43_virusinfo = (int)fVirusPrompted.SetBoolean(field_43_virusinfo, value);


        }

        /**
         * 
         * @return  the fVirusPrompted field value.
         */
        public bool IsFVirusPrompted()
        {
            return fVirusPrompted.IsSet(field_43_virusinfo);

        }

        /**
         * Sets the fVirusLoadSafe field value.
         * 
         */
        public void SetFVirusLoadSafe(bool value)
        {
            field_43_virusinfo = (int)fVirusLoadSafe.SetBoolean(field_43_virusinfo, value);


        }

        /**
         * 
         * @return  the fVirusLoadSafe field value.
         */
        public bool IsFVirusLoadSafe()
        {
            return fVirusLoadSafe.IsSet(field_43_virusinfo);

        }

        /**
         * Sets the KeyVirusSession30 field value.
         * 
         */
        public void SetKeyVirusSession30(int value)
        {
            field_43_virusinfo = (int)KeyVirusSession30.SetValue(field_43_virusinfo, value);


        }

        /**
         * 
         * @return  the KeyVirusSession30 field value.
         */
        public int GetKeyVirusSession30()
        {
            return (int)KeyVirusSession30.GetValue(field_43_virusinfo);

        }


    }
}



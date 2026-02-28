
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

    /**
     * Character Properties.
     * NOTE: This source Is automatically generated please do not modify this file.  Either subclass or
     *       remove the record in src/records/definitions.

     * @author S. Ryan Ackley
     */
    public abstract class CHPAbstractType : BaseObject
    {

        protected short field_1_chse;
        protected int field_2_format_flags;
        private static BitField fBold = BitFieldFactory.GetInstance(0x0001);
        private static BitField fItalic = BitFieldFactory.GetInstance(0x0002);
        private static BitField fRMarkDel = BitFieldFactory.GetInstance(0x0004);
        private static BitField fOutline = BitFieldFactory.GetInstance(0x0008);
        private static BitField fFldVanish = BitFieldFactory.GetInstance(0x0010);
        private static BitField fSmallCaps = BitFieldFactory.GetInstance(0x0020);
        private static BitField fCaps = BitFieldFactory.GetInstance(0x0040);
        private static BitField fVanish = BitFieldFactory.GetInstance(0x0080);
        private static BitField fRMark = BitFieldFactory.GetInstance(0x0100);
        private static BitField fSpec = BitFieldFactory.GetInstance(0x0200);
        private static BitField fStrike = BitFieldFactory.GetInstance(0x0400);
        private static BitField fObj = BitFieldFactory.GetInstance(0x0800);
        private static BitField fShadow = BitFieldFactory.GetInstance(0x1000);
        private static BitField fLowerCase = BitFieldFactory.GetInstance(0x2000);
        private static BitField fData = BitFieldFactory.GetInstance(0x4000);
        private static BitField fOle2 = BitFieldFactory.GetInstance(0x8000);
        protected int field_3_format_flags1;
        private static BitField fEmboss = BitFieldFactory.GetInstance(0x0001);
        private static BitField fImprint = BitFieldFactory.GetInstance(0x0002);
        private static BitField fDStrike = BitFieldFactory.GetInstance(0x0004);
        private static BitField fUsePgsuSettings = BitFieldFactory.GetInstance(0x0008);
        protected int field_4_ftcAscii;
        protected int field_5_ftcFE;
        protected int field_6_ftcOther;
        protected int field_7_hps;
        protected int field_8_dxaSpace;
        protected byte field_9_iss;
        protected byte field_10_kul;
        protected byte field_11_ico;
        protected int field_12_hpsPos;
        protected int field_13_lidDefault;
        protected int field_14_lidFE;
        protected byte field_15_idctHint;
        protected int field_16_wCharScale;
        protected int field_17_fcPic;
        protected int field_18_fcObj;
        protected int field_19_lTagObj;
        protected int field_20_ibstRMark;
        protected int field_21_ibstRMarkDel;
        protected DateAndTime field_22_dttmRMark;
        protected DateAndTime field_23_dttmRMarkDel;
        protected int field_24_istd;
        protected int field_25_baseIstd;
        protected int field_26_ftcSym;
        protected int field_27_xchSym;
        protected int field_28_idslRMReason;
        protected int field_29_idslReasonDel;
        protected byte field_30_ysr;
        protected byte field_31_chYsr;
        protected int field_32_hpsKern;
        protected short field_33_Highlight;
        private static BitField icoHighlight = BitFieldFactory.GetInstance(0x001f);
        private static BitField fHighlight = BitFieldFactory.GetInstance(0x0020);
        private static BitField kcd = BitFieldFactory.GetInstance(0x01c0);
        private static BitField fNavHighlight = BitFieldFactory.GetInstance(0x0200);
        private static BitField fChsDiff = BitFieldFactory.GetInstance(0x0400);
        private static BitField fMacChs = BitFieldFactory.GetInstance(0x0800);
        private static BitField fFtcAsciSym = BitFieldFactory.GetInstance(0x1000);
        protected short field_34_fPropMark;
        protected int field_35_ibstPropRMark;
        protected DateAndTime field_36_dttmPropRMark;
        protected byte field_37_sfxtText;
        protected byte field_38_fDispFldRMark;
        protected int field_39_ibstDispFldRMark;
        protected DateAndTime field_40_dttmDispFldRMark;
        protected byte[] field_41_xstDispFldRMark;
        protected ShadingDescriptor field_42_shd;
        protected BorderCode field_43_brc;


        public CHPAbstractType()
        {

        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public int GetSize()
        {
            return 4 + +2 + 2 + 2 + 2 + 2 + 2 + 2 + 4 + 1 + 1 + 1 + 2 + 2 + 2 + 1 + 2 + 4 + 4 + 4 + 2 + 2 + 4 + 4 + 2 + 2 + 2 + 2 + 2 + 2 + 1 + 1 + 2 + 2 + 2 + 2 + 4 + 1 + 1 + 2 + 4 + 32 + 2 + 4;
        }



        /**
         * Get the chse field for the CHP record.
         */
        public short GetChse()
        {
            return field_1_chse;
        }

        /**
         * Set the chse field for the CHP record.
         */
        public void SetChse(short field_1_chse)
        {
            this.field_1_chse = field_1_chse;
        }

        /**
         * Get the format_flags field for the CHP record.
         */
        public int GetFormat_flags()
        {
            return field_2_format_flags;
        }

        /**
         * Set the format_flags field for the CHP record.
         */
        public void SetFormat_flags(int field_2_format_flags)
        {
            this.field_2_format_flags = field_2_format_flags;
        }

        /**
         * Get the format_flags1 field for the CHP record.
         */
        public int GetFormat_flags1()
        {
            return field_3_format_flags1;
        }

        /**
         * Set the format_flags1 field for the CHP record.
         */
        public void SetFormat_flags1(int field_3_format_flags1)
        {
            this.field_3_format_flags1 = field_3_format_flags1;
        }

        /**
         * Get the ftcAscii field for the CHP record.
         */
        public int GetFtcAscii()
        {
            return field_4_ftcAscii;
        }

        /**
         * Set the ftcAscii field for the CHP record.
         */
        public void SetFtcAscii(int field_4_ftcAscii)
        {
            this.field_4_ftcAscii = field_4_ftcAscii;
        }

        /**
         * Get the ftcFE field for the CHP record.
         */
        public int GetFtcFE()
        {
            return field_5_ftcFE;
        }

        /**
         * Set the ftcFE field for the CHP record.
         */
        public void SetFtcFE(int field_5_ftcFE)
        {
            this.field_5_ftcFE = field_5_ftcFE;
        }

        /**
         * Get the ftcOther field for the CHP record.
         */
        public int GetFtcOther()
        {
            return field_6_ftcOther;
        }

        /**
         * Set the ftcOther field for the CHP record.
         */
        public void SetFtcOther(int field_6_ftcOther)
        {
            this.field_6_ftcOther = field_6_ftcOther;
        }

        /**
         * Get the hps field for the CHP record.
         */
        public int GetHps()
        {
            return field_7_hps;
        }

        /**
         * Set the hps field for the CHP record.
         */
        public void SetHps(int field_7_hps)
        {
            this.field_7_hps = field_7_hps;
        }

        /**
         * Get the dxaSpace field for the CHP record.
         */
        public int GetDxaSpace()
        {
            return field_8_dxaSpace;
        }

        /**
         * Set the dxaSpace field for the CHP record.
         */
        public void SetDxaSpace(int field_8_dxaSpace)
        {
            this.field_8_dxaSpace = field_8_dxaSpace;
        }

        /**
         * Get the Iss field for the CHP record.
         */
        public byte GetIss()
        {
            return field_9_iss;
        }

        /**
         * Set the Iss field for the CHP record.
         */
        public void SetIss(byte field_9_iss)
        {
            this.field_9_iss = field_9_iss;
        }

        /**
         * Get the kul field for the CHP record.
         */
        public byte GetKul()
        {
            return field_10_kul;
        }

        /**
         * Set the kul field for the CHP record.
         */
        public void SetKul(byte field_10_kul)
        {
            this.field_10_kul = field_10_kul;
        }

        /**
         * Get the ico field for the CHP record.
         */
        public byte GetIco()
        {
            return field_11_ico;
        }

        /**
         * Set the ico field for the CHP record.
         */
        public void SetIco(byte field_11_ico)
        {
            this.field_11_ico = field_11_ico;
        }

        /**
         * Get the hpsPos field for the CHP record.
         */
        public int GetHpsPos()
        {
            return field_12_hpsPos;
        }

        /**
         * Set the hpsPos field for the CHP record.
         */
        public void SetHpsPos(int field_12_hpsPos)
        {
            this.field_12_hpsPos = field_12_hpsPos;
        }

        /**
         * Get the lidDefault field for the CHP record.
         */
        public int GetLidDefault()
        {
            return field_13_lidDefault;
        }

        /**
         * Set the lidDefault field for the CHP record.
         */
        public void SetLidDefault(int field_13_lidDefault)
        {
            this.field_13_lidDefault = field_13_lidDefault;
        }

        /**
         * Get the lidFE field for the CHP record.
         */
        public int GetLidFE()
        {
            return field_14_lidFE;
        }

        /**
         * Set the lidFE field for the CHP record.
         */
        public void SetLidFE(int field_14_lidFE)
        {
            this.field_14_lidFE = field_14_lidFE;
        }

        /**
         * Get the idctHint field for the CHP record.
         */
        public byte GetIdctHint()
        {
            return field_15_idctHint;
        }

        /**
         * Set the idctHint field for the CHP record.
         */
        public void SetIdctHint(byte field_15_idctHint)
        {
            this.field_15_idctHint = field_15_idctHint;
        }

        /**
         * Get the wCharScale field for the CHP record.
         */
        public int GetWCharScale()
        {
            return field_16_wCharScale;
        }

        /**
         * Set the wCharScale field for the CHP record.
         */
        public void SetWCharScale(int field_16_wCharScale)
        {
            this.field_16_wCharScale = field_16_wCharScale;
        }

        /**
         * Get the fcPic field for the CHP record.
         */
        public int GetFcPic()
        {
            return field_17_fcPic;
        }

        /**
         * Set the fcPic field for the CHP record.
         */
        public void SetFcPic(int field_17_fcPic)
        {
            this.field_17_fcPic = field_17_fcPic;
        }

        /**
         * Get the fcObj field for the CHP record.
         */
        public int GetFcObj()
        {
            return field_18_fcObj;
        }

        /**
         * Set the fcObj field for the CHP record.
         */
        public void SetFcObj(int field_18_fcObj)
        {
            this.field_18_fcObj = field_18_fcObj;
        }

        /**
         * Get the lTagObj field for the CHP record.
         */
        public int GetLTagObj()
        {
            return field_19_lTagObj;
        }

        /**
         * Set the lTagObj field for the CHP record.
         */
        public void SetLTagObj(int field_19_lTagObj)
        {
            this.field_19_lTagObj = field_19_lTagObj;
        }

        /**
         * Get the ibstRMark field for the CHP record.
         */
        public int GetIbstRMark()
        {
            return field_20_ibstRMark;
        }

        /**
         * Set the ibstRMark field for the CHP record.
         */
        public void SetIbstRMark(int field_20_ibstRMark)
        {
            this.field_20_ibstRMark = field_20_ibstRMark;
        }

        /**
         * Get the ibstRMarkDel field for the CHP record.
         */
        public int GetIbstRMarkDel()
        {
            return field_21_ibstRMarkDel;
        }

        /**
         * Set the ibstRMarkDel field for the CHP record.
         */
        public void SetIbstRMarkDel(int field_21_ibstRMarkDel)
        {
            this.field_21_ibstRMarkDel = field_21_ibstRMarkDel;
        }

        /**
         * Get the dttmRMark field for the CHP record.
         */
        public DateAndTime GetDttmRMark()
        {
            return field_22_dttmRMark;
        }

        /**
         * Set the dttmRMark field for the CHP record.
         */
        public void SetDttmRMark(DateAndTime field_22_dttmRMark)
        {
            this.field_22_dttmRMark = field_22_dttmRMark;
        }

        /**
         * Get the dttmRMarkDel field for the CHP record.
         */
        public DateAndTime GetDttmRMarkDel()
        {
            return field_23_dttmRMarkDel;
        }

        /**
         * Set the dttmRMarkDel field for the CHP record.
         */
        public void SetDttmRMarkDel(DateAndTime field_23_dttmRMarkDel)
        {
            this.field_23_dttmRMarkDel = field_23_dttmRMarkDel;
        }

        /**
         * Get the Istd field for the CHP record.
         */
        public int GetIstd()
        {
            return field_24_istd;
        }

        /**
         * Set the Istd field for the CHP record.
         */
        public void SetIstd(int field_24_istd)
        {
            this.field_24_istd = field_24_istd;
        }

        /**
         * Get the baseIstd field for the CHP record.
         */
        public int GetBaseIstd()
        {
            return field_25_baseIstd;
        }

        /**
         * Set the baseIstd field for the CHP record.
         */
        public void SetBaseIstd(int field_25_baseIstd)
        {
            this.field_25_baseIstd = field_25_baseIstd;
        }

        /**
         * Get the ftcSym field for the CHP record.
         */
        public int GetFtcSym()
        {
            return field_26_ftcSym;
        }

        /**
         * Set the ftcSym field for the CHP record.
         */
        public void SetFtcSym(int field_26_ftcSym)
        {
            this.field_26_ftcSym = field_26_ftcSym;
        }

        /**
         * Get the xchSym field for the CHP record.
         */
        public int GetXchSym()
        {
            return field_27_xchSym;
        }

        /**
         * Set the xchSym field for the CHP record.
         */
        public void SetXchSym(int field_27_xchSym)
        {
            this.field_27_xchSym = field_27_xchSym;
        }

        /**
         * Get the idslRMReason field for the CHP record.
         */
        public int GetIdslRMReason()
        {
            return field_28_idslRMReason;
        }

        /**
         * Set the idslRMReason field for the CHP record.
         */
        public void SetIdslRMReason(int field_28_idslRMReason)
        {
            this.field_28_idslRMReason = field_28_idslRMReason;
        }

        /**
         * Get the idslReasonDel field for the CHP record.
         */
        public int GetIdslReasonDel()
        {
            return field_29_idslReasonDel;
        }

        /**
         * Set the idslReasonDel field for the CHP record.
         */
        public void SetIdslReasonDel(int field_29_idslReasonDel)
        {
            this.field_29_idslReasonDel = field_29_idslReasonDel;
        }

        /**
         * Get the ysr field for the CHP record.
         */
        public byte GetYsr()
        {
            return field_30_ysr;
        }

        /**
         * Set the ysr field for the CHP record.
         */
        public void SetYsr(byte field_30_ysr)
        {
            this.field_30_ysr = field_30_ysr;
        }

        /**
         * Get the chYsr field for the CHP record.
         */
        public byte GetChYsr()
        {
            return field_31_chYsr;
        }

        /**
         * Set the chYsr field for the CHP record.
         */
        public void SetChYsr(byte field_31_chYsr)
        {
            this.field_31_chYsr = field_31_chYsr;
        }

        /**
         * Get the hpsKern field for the CHP record.
         */
        public int GetHpsKern()
        {
            return field_32_hpsKern;
        }

        /**
         * Set the hpsKern field for the CHP record.
         */
        public void SetHpsKern(int field_32_hpsKern)
        {
            this.field_32_hpsKern = field_32_hpsKern;
        }

        /**
         * Get the Highlight field for the CHP record.
         */
        public short GetHighlight()
        {
            return field_33_Highlight;
        }

        /**
         * Set the Highlight field for the CHP record.
         */
        public void SetHighlight(short field_33_Highlight)
        {
            this.field_33_Highlight = field_33_Highlight;
        }

        /**
         * Get the fPropMark field for the CHP record.
         */
        public short GetFPropMark()
        {
            return field_34_fPropMark;
        }

        /**
         * Set the fPropMark field for the CHP record.
         */
        public void SetFPropMark(short field_34_fPropMark)
        {
            this.field_34_fPropMark = field_34_fPropMark;
        }

        /**
         * Get the ibstPropRMark field for the CHP record.
         */
        public int GetIbstPropRMark()
        {
            return field_35_ibstPropRMark;
        }

        /**
         * Set the ibstPropRMark field for the CHP record.
         */
        public void SetIbstPropRMark(int field_35_ibstPropRMark)
        {
            this.field_35_ibstPropRMark = field_35_ibstPropRMark;
        }

        /**
         * Get the dttmPropRMark field for the CHP record.
         */
        public DateAndTime GetDttmPropRMark()
        {
            return field_36_dttmPropRMark;
        }

        /**
         * Set the dttmPropRMark field for the CHP record.
         */
        public void SetDttmPropRMark(DateAndTime field_36_dttmPropRMark)
        {
            this.field_36_dttmPropRMark = field_36_dttmPropRMark;
        }

        /**
         * Get the sfxtText field for the CHP record.
         */
        public byte GetSfxtText()
        {
            return field_37_sfxtText;
        }

        /**
         * Set the sfxtText field for the CHP record.
         */
        public void SetSfxtText(byte field_37_sfxtText)
        {
            this.field_37_sfxtText = field_37_sfxtText;
        }

        /**
         * Get the fDispFldRMark field for the CHP record.
         */
        public byte GetFDispFldRMark()
        {
            return field_38_fDispFldRMark;
        }

        /**
         * Set the fDispFldRMark field for the CHP record.
         */
        public void SetFDispFldRMark(byte field_38_fDispFldRMark)
        {
            this.field_38_fDispFldRMark = field_38_fDispFldRMark;
        }

        /**
         * Get the ibstDispFldRMark field for the CHP record.
         */
        public int GetIbstDispFldRMark()
        {
            return field_39_ibstDispFldRMark;
        }

        /**
         * Set the ibstDispFldRMark field for the CHP record.
         */
        public void SetIbstDispFldRMark(int field_39_ibstDispFldRMark)
        {
            this.field_39_ibstDispFldRMark = field_39_ibstDispFldRMark;
        }

        /**
         * Get the dttmDispFldRMark field for the CHP record.
         */
        public DateAndTime GetDttmDispFldRMark()
        {
            return field_40_dttmDispFldRMark;
        }

        /**
         * Set the dttmDispFldRMark field for the CHP record.
         */
        public void SetDttmDispFldRMark(DateAndTime field_40_dttmDispFldRMark)
        {
            this.field_40_dttmDispFldRMark = field_40_dttmDispFldRMark;
        }

        /**
         * Get the xstDispFldRMark field for the CHP record.
         */
        public byte[] GetXstDispFldRMark()
        {
            return field_41_xstDispFldRMark;
        }

        /**
         * Set the xstDispFldRMark field for the CHP record.
         */
        public void SetXstDispFldRMark(byte[] field_41_xstDispFldRMark)
        {
            this.field_41_xstDispFldRMark = field_41_xstDispFldRMark;
        }

        /**
         * Get the shd field for the CHP record.
         */
        public ShadingDescriptor GetShd()
        {
            return field_42_shd;
        }

        /**
         * Set the shd field for the CHP record.
         */
        public void SetShd(ShadingDescriptor field_42_shd)
        {
            this.field_42_shd = field_42_shd;
        }

        /**
         * Get the brc field for the CHP record.
         */
        public BorderCode GetBrc()
        {
            return field_43_brc;
        }

        /**
         * Set the brc field for the CHP record.
         */
        public void SetBrc(BorderCode field_43_brc)
        {
            this.field_43_brc = field_43_brc;
        }

        /**
         * Sets the fBold field value.
         * 
         */
        public void SetFBold(bool value)
        {
            field_2_format_flags = (int)fBold.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fBold field value.
         */
        public bool IsFBold()
        {
            return fBold.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fItalic field value.
         * 
         */
        public void SetFItalic(bool value)
        {
            field_2_format_flags = (int)fItalic.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fItalic field value.
         */
        public bool IsFItalic()
        {
            return fItalic.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fRMarkDel field value.
         * 
         */
        public void SetFRMarkDel(bool value)
        {
            field_2_format_flags = (int)fRMarkDel.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fRMarkDel field value.
         */
        public bool IsFRMarkDel()
        {
            return fRMarkDel.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fOutline field value.
         * 
         */
        public void SetFOutline(bool value)
        {
            field_2_format_flags = (int)fOutline.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fOutline field value.
         */
        public bool IsFOutline()
        {
            return fOutline.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fFldVanish field value.
         * 
         */
        public void SetFFldVanish(bool value)
        {
            field_2_format_flags = (int)fFldVanish.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fFldVanish field value.
         */
        public bool IsFFldVanish()
        {
            return fFldVanish.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fSmallCaps field value.
         * 
         */
        public void SetFSmallCaps(bool value)
        {
            field_2_format_flags = (int)fSmallCaps.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fSmallCaps field value.
         */
        public bool IsFSmallCaps()
        {
            return fSmallCaps.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fCaps field value.
         * 
         */
        public void SetFCaps(bool value)
        {
            field_2_format_flags = (int)fCaps.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fCaps field value.
         */
        public bool IsFCaps()
        {
            return fCaps.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fVanish field value.
         * 
         */
        public void SetFVanish(bool value)
        {
            field_2_format_flags = (int)fVanish.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fVanish field value.
         */
        public bool IsFVanish()
        {
            return fVanish.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fRMark field value.
         * 
         */
        public void SetFRMark(bool value)
        {
            field_2_format_flags = (int)fRMark.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fRMark field value.
         */
        public bool IsFRMark()
        {
            return fRMark.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fSpec field value.
         * 
         */
        public void SetFSpec(bool value)
        {
            field_2_format_flags = (int)fSpec.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fSpec field value.
         */
        public bool IsFSpec()
        {
            return fSpec.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fStrike field value.
         * 
         */
        public void SetFStrike(bool value)
        {
            field_2_format_flags = (int)fStrike.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fStrike field value.
         */
        public bool IsFStrike()
        {
            return fStrike.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fObj field value.
         * 
         */
        public void SetFObj(bool value)
        {
            field_2_format_flags = (int)fObj.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fObj field value.
         */
        public bool IsFObj()
        {
            return fObj.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fShadow field value.
         * 
         */
        public void SetFShadow(bool value)
        {
            field_2_format_flags = (int)fShadow.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fShadow field value.
         */
        public bool IsFShadow()
        {
            return fShadow.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fLowerCase field value.
         * 
         */
        public void SetFLowerCase(bool value)
        {
            field_2_format_flags = (int)fLowerCase.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fLowerCase field value.
         */
        public bool IsFLowerCase()
        {
            return fLowerCase.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fData field value.
         * 
         */
        public void SetFData(bool value)
        {
            field_2_format_flags = (int)fData.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fData field value.
         */
        public bool IsFData()
        {
            return fData.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fOle2 field value.
         * 
         */
        public void SetFOle2(bool value)
        {
            field_2_format_flags = (int)fOle2.SetBoolean(field_2_format_flags, value);


        }

        /**
         * 
         * @return  the fOle2 field value.
         */
        public bool IsFOle2()
        {
            return fOle2.IsSet(field_2_format_flags);

        }

        /**
         * Sets the fEmboss field value.
         * 
         */
        public void SetFEmboss(bool value)
        {
            field_3_format_flags1 = (int)fEmboss.SetBoolean(field_3_format_flags1, value);


        }

        /**
         * 
         * @return  the fEmboss field value.
         */
        public bool IsFEmboss()
        {
            return fEmboss.IsSet(field_3_format_flags1);

        }

        /**
         * Sets the fImprint field value.
         * 
         */
        public void SetFImprint(bool value)
        {
            field_3_format_flags1 = (int)fImprint.SetBoolean(field_3_format_flags1, value);


        }

        /**
         * 
         * @return  the fImprint field value.
         */
        public bool IsFImprint()
        {
            return fImprint.IsSet(field_3_format_flags1);

        }

        /**
         * Sets the fDStrike field value.
         * 
         */
        public void SetFDStrike(bool value)
        {
            field_3_format_flags1 = (int)fDStrike.SetBoolean(field_3_format_flags1, value);


        }

        /**
         * 
         * @return  the fDStrike field value.
         */
        public bool IsFDStrike()
        {
            return fDStrike.IsSet(field_3_format_flags1);

        }

        /**
         * Sets the fUsePgsuSettings field value.
         * 
         */
        public void SetFUsePgsuSettings(bool value)
        {
            field_3_format_flags1 = (int)fUsePgsuSettings.SetBoolean(field_3_format_flags1, value);


        }

        /**
         * 
         * @return  the fUsePgsuSettings field value.
         */
        public bool IsFUsePgsuSettings()
        {
            return fUsePgsuSettings.IsSet(field_3_format_flags1);

        }

        /**
         * Sets the icoHighlight field value.
         * 
         */
        public void SetIcoHighlight(byte value)
        {
            field_33_Highlight = (short)icoHighlight.SetValue(field_33_Highlight, value);


        }

        /**
         * 
         * @return  the icoHighlight field value.
         */
        public byte GetIcoHighlight()
        {
            return (byte)icoHighlight.GetValue(field_33_Highlight);

        }

        /**
         * Sets the fHighlight field value.
         * 
         */
        public void SetFHighlight(bool value)
        {
            field_33_Highlight = (short)fHighlight.SetBoolean(field_33_Highlight, value);


        }

        /**
         * 
         * @return  the fHighlight field value.
         */
        public bool IsFHighlight()
        {
            return fHighlight.IsSet(field_33_Highlight);

        }

        /**
         * Sets the kcd field value.
         * 
         */
        public void SetKcd(byte value)
        {
            field_33_Highlight = (short)kcd.SetValue(field_33_Highlight, value);


        }

        /**
         * 
         * @return  the kcd field value.
         */
        public byte GetKcd()
        {
            return (byte)kcd.GetValue(field_33_Highlight);

        }

        /**
         * Sets the fNavHighlight field value.
         * 
         */
        public void SetFNavHighlight(bool value)
        {
            field_33_Highlight = (short)fNavHighlight.SetBoolean(field_33_Highlight, value);


        }

        /**
         * 
         * @return  the fNavHighlight field value.
         */
        public bool IsFNavHighlight()
        {
            return fNavHighlight.IsSet(field_33_Highlight);

        }

        /**
         * Sets the fChsDiff field value.
         * 
         */
        public void SetFChsDiff(bool value)
        {
            field_33_Highlight = (short)fChsDiff.SetBoolean(field_33_Highlight, value);


        }

        /**
         * 
         * @return  the fChsDiff field value.
         */
        public bool IsFChsDiff()
        {
            return fChsDiff.IsSet(field_33_Highlight);

        }

        /**
         * Sets the fMacChs field value.
         * 
         */
        public void SetFMacChs(bool value)
        {
            field_33_Highlight = (short)fMacChs.SetBoolean(field_33_Highlight, value);


        }

        /**
         * 
         * @return  the fMacChs field value.
         */
        public bool IsFMacChs()
        {
            return fMacChs.IsSet(field_33_Highlight);

        }

        /**
         * Sets the fFtcAsciSym field value.
         * 
         */
        public void SetFFtcAsciSym(bool value)
        {
            field_33_Highlight = (short)fFtcAsciSym.SetBoolean(field_33_Highlight, value);


        }

        /**
         * 
         * @return  the fFtcAsciSym field value.
         */
        public bool IsFFtcAsciSym()
        {
            return fFtcAsciSym.IsSet(field_33_Highlight);

        }


    }
}

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


namespace NPOI.HWPF.UserModel
{
    using System;
    using NPOI.HWPF.Model.Types;

    /**
     * @author Ryan Ackley
     */
    public class CharacterProperties
      : CHPAbstractType //, ICloneable
    {
        public static short SPRM_FRMARKDEL = (short)0x0800;
        public static short SPRM_FRMARK = 0x0801;
        public static short SPRM_FFLDVANISH = 0x0802;
        public static short SPRM_PICLOCATION = 0x6A03;
        public static short SPRM_IBSTRMARK = 0x4804;
        public static short SPRM_DTTMRMARK = 0x6805;
        public static short SPRM_FDATA = 0x0806;
        public static short SPRM_SYMBOL = 0x6A09;
        public static short SPRM_FOLE2 = 0x080A;
        public static short SPRM_HIGHLIGHT = 0x2A0C;
        public static short SPRM_OBJLOCATION = 0x680E;
        public static short SPRM_ISTD = 0x4A30;
        public static short SPRM_FBOLD = 0x0835;
        public static short SPRM_FITALIC = 0x0836;
        public static short SPRM_FSTRIKE = 0x0837;
        public static short SPRM_FOUTLINE = 0x0838;
        public static short SPRM_FSHADOW = 0x0839;
        public static short SPRM_FSMALLCAPS = 0x083A;
        public static short SPRM_FCAPS = 0x083B;
        public static short SPRM_FVANISH = 0x083C;
        public static short SPRM_KUL = 0x2A3E;
        public static short SPRM_DXASPACE = unchecked((short)0x8840);
        public static short SPRM_LID = 0x4A41;
        public static short SPRM_ICO = 0x2A42;
        public static short SPRM_HPS = 0x4A43;
        public static short SPRM_HPSPOS = 0x4845;
        public static short SPRM_ISS = 0x2A48;
        public static short SPRM_HPSKERN = 0x484B;
        public static short SPRM_YSRI = 0x484E;
        public static short SPRM_RGFTCASCII = 0x4A4F;
        public static short SPRM_RGFTCFAREAST = 0x4A50;
        public static short SPRM_RGFTCNOTFAREAST = 0x4A51;
        public static short SPRM_CHARSCALE = 0x4852;
        public static short SPRM_FDSTRIKE = 0x2A53;
        public static short SPRM_FIMPRINT = 0x0854;
        public static short SPRM_FSPEC = 0x0855;
        public static short SPRM_FOBJ = 0x0856;
        public static short SPRM_PROPRMARK = unchecked((short)0xCA57);
        public static short SPRM_FEMBOSS = 0x0858;
        public static short SPRM_SFXTEXT = 0x2859;
        public static short SPRM_DISPFLDRMARK = unchecked((short)0xCA62);
        public static short SPRM_IBSTRMARKDEL = 0x4863;
        public static short SPRM_DTTMRMARKDEL = 0x6864;
        public static short SPRM_BRC = 0x6865;
        public static short SPRM_SHD = 0x4866;
        public static short SPRM_IDSIRMARKDEL = 0x4867;
        public static short SPRM_CPG = 0x486B;
        public static short SPRM_NONFELID = 0x486D;
        public static short SPRM_FELID = 0x486E;
        public static short SPRM_IDCTHINT = 0x286F;

        int _ico24 = -1; // default to -1 so we can ignore it for word 97 files

        public CharacterProperties()
        {
            field_17_fcPic = -1;
            field_22_dttmRMark = new DateAndTime();
            field_23_dttmRMarkDel = new DateAndTime();
            field_36_dttmPropRMark = new DateAndTime();
            field_40_dttmDispFldRMark = new DateAndTime();
            field_41_xstDispFldRMark = new byte[36];
            field_42_shd = new ShadingDescriptor();
            field_43_brc = new BorderCode();
            field_7_hps = 20;
            field_24_istd = 10;
            field_16_wCharScale = 100;
            field_13_lidDefault = 0x0400;
            field_14_lidFE = 0x0400;
        }

        public bool IsMarkedDeleted()
        {
            return IsFRMarkDel();
        }

        public void markDeleted(bool mark)
        {
            base.SetFRMarkDel(mark);
        }

        public bool IsBold()
        {
            return IsFBold();
        }

        public void SetBold(bool bold)
        {
            base.SetFBold(bold);
        }

        public bool IsItalic()
        {
            return IsFItalic();
        }

        public void SetItalic(bool italic)
        {
            base.SetFItalic(italic);
        }

        public bool IsOutlined()
        {
            return IsFOutline();
        }

        public void SetOutline(bool outlined)
        {
            base.SetFOutline(outlined);
        }

        public bool IsFldVanished()
        {
            return IsFFldVanish();
        }

        public void SetFldVanish(bool fldVanish)
        {
            base.SetFFldVanish(fldVanish);
        }

        public bool IsSmallCaps()
        {
            return IsFSmallCaps();
        }

        public void SetSmallCaps(bool smallCaps)
        {
            base.SetFSmallCaps(smallCaps);
        }

        public bool IsCapitalized()
        {
            return IsFCaps();
        }

        public void SetCapitalized(bool caps)
        {
            base.SetFCaps(caps);
        }

        public bool IsVanished()
        {
            return IsFVanish();
        }

        public void SetVanished(bool vanish)
        {
            base.SetFVanish(vanish);

        }
        public bool IsMarkedInserted()
        {
            return IsFRMark();
        }

        public void markInserted(bool mark)
        {
            base.SetFRMark(mark);
        }

        public bool IsStrikeThrough()
        {
            return IsFStrike();
        }

        public void strikeThrough(bool strike)
        {
            base.SetFStrike(strike);
        }

        public bool IsShadowed()
        {
            return IsFShadow();
        }

        public void SetShadow(bool shadow)
        {
            base.SetFShadow(shadow);

        }

        public bool IsEmbossed()
        {
            return IsFEmboss();
        }

        public void SetEmbossed(bool emboss)
        {
            base.SetFEmboss(emboss);
        }

        public bool IsImprinted()
        {
            return IsFImprint();
        }

        public void SetImprinted(bool imprint)
        {
            base.SetFImprint(imprint);
        }

        public bool IsDoubleStrikeThrough()
        {
            return IsFDStrike();
        }

        public void SetDoubleStrikeThrough(bool dstrike)
        {
            base.SetFDStrike(dstrike);
        }

        public int GetFontSize()
        {
            return GetHps();
        }

        public void SetFontSize(int halfPoints)
        {
            base.SetHps(halfPoints);
        }

        public int GetCharacterSpacing()
        {
            return GetDxaSpace();
        }

        public void SetCharacterSpacing(int twips)
        {
            base.SetDxaSpace(twips);
        }

        public short GetSubSuperScriptIndex()
        {
            return GetIss();
        }

        public void SetSubSuperScriptIndex(short Iss)
        {
            base.SetDxaSpace(Iss);
        }

        public int GetUnderlineCode()
        {
            return base.GetKul();
        }

        public void SetUnderlineCode(int kul)
        {
            base.SetKul((byte)kul);
        }

        public int GetColor()
        {
            return base.GetIco();
        }

        public void SetColor(int color)
        {
            base.SetIco((byte)color);
        }

        public int GetVerticalOffset()
        {
            return base.GetHpsPos();
        }

        public void SetVerticalOffset(int hpsPos)
        {
            base.SetHpsPos(hpsPos);
        }

        public int GetKerning()
        {
            return base.GetHpsKern();
        }

        public void SetKerning(int kern)
        {
            base.SetHpsKern(kern);
        }

        public bool IsHighlighted()
        {
            return base.IsFHighlight();
        }

        public void SetHighlighted(byte color)
        {
            base.SetIcoHighlight(color);
        }

        /**
        * Get the ico24 field for the CHP record.
        */
        public int GetIco24()
        {
            if (_ico24 == -1)
            {
                switch (field_11_ico) // convert word 97 colour numbers to 0xBBGGRR value
                {
                    case 0: // auto
                        return -1;
                    case 1: // black
                        return 0x000000;
                    case 2: // blue
                        return 0xFF0000;
                    case 3: // cyan
                        return 0xFFFF00;
                    case 4: // green
                        return 0x00FF00;
                    case 5: // magenta
                        return 0xFF00FF;
                    case 6: // red
                        return 0x0000FF;
                    case 7: // yellow
                        return 0x00FFFF;
                    case 8: // white
                        return 0x0FFFFFF;
                    case 9: // dark blue
                        return 0x800000;
                    case 10: // dark cyan
                        return 0x808000;
                    case 11: // dark green
                        return 0x008000;
                    case 12: // dark magenta
                        return 0x800080;
                    case 13: // dark red
                        return 0x000080;
                    case 14: // dark yellow
                        return 0x008080;
                    case 15: // dark grey
                        return 0x808080;
                    case 16: // light grey
                        return 0xC0C0C0;
                }
            }
            return _ico24;
        }

        /**
         * Set the ico24 field for the CHP record.
         */
        public void SetIco24(int colour24)
        {
            _ico24 = colour24 & 0xFFFFFF; // only keep the 24bit 0xBBGGRR colour
        }

        //public Object Clone()
        //{
        //    CharacterProperties cp = (CharacterProperties)base.Clone();
        //    cp.field_22_dttmRMark = (DateAndTime)field_22_dttmRMark.Clone();
        //    cp.field_23_dttmRMarkDel = (DateAndTime)field_23_dttmRMarkDel.Clone();
        //    cp.field_36_dttmPropRMark = (DateAndTime)field_36_dttmPropRMark.Clone();
        //    cp.field_40_dttmDispFldRMark = (DateAndTime)field_40_dttmDispFldRMark.Clone();
        //    cp.field_41_xstDispFldRMark = (byte[])field_41_xstDispFldRMark.Clone();
        //    cp.field_42_shd = (ShadingDescriptor)field_42_shd.Clone();

        //    cp._ico24 = _ico24;

        //    return cp;
        //}
    }
}
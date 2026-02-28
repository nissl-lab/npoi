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

namespace NPOI.HWPF.SPRM
{

    using NPOI.HWPF.UserModel;
    using NPOI.Util;
    using System;

    public class CharacterSprmUncompressor
    {
        public CharacterSprmUncompressor()
        {
        }

        public static CharacterProperties UncompressCHP(CharacterProperties parent,
                                                        byte[] grpprl,
                                                        int Offset)
        {
            CharacterProperties newProperties = null;
            newProperties = (CharacterProperties)parent.Clone();

            SprmIterator sprmIt = new SprmIterator(grpprl, Offset);

            while (sprmIt.HasNext())
            {
                SprmOperation sprm = sprmIt.Next();
                UncompressCHPOperation(parent, newProperties, sprm);
            }

            return newProperties;
        }

        /**
         * Used in decompression of a chpx. This performs an operation defined by
         * a single sprm.
         *
         * @param oldCHP The base CharacterProperties.
         * @param newCHP The current CharacterProperties.
         * @param operand The operand defined by the sprm (See Word file format spec)
         * @param param The parameter defined by the sprm (See Word file format spec)
         * @param varParam The variable length parameter defined by the sprm. (See
         *        Word file format spec)
         * @param grpprl The entire chpx that this operation is a part of.
         * @param offset The offset in the grpprl of the next sprm
         * @param styleSheet The StyleSheet for this document.
         */
        static void UncompressCHPOperation(CharacterProperties oldCHP,
                                            CharacterProperties newCHP,
                                            SprmOperation sprm)
        {

            switch (sprm.Operation)
            {
                case 0:
                    newCHP.SetFRMarkDel(GetFlag(sprm.Operand));
                    break;
                case 0x1:
                    newCHP.SetFRMark(GetFlag(sprm.Operand));
                    break;
                case 0x2:
                    newCHP.SetFFldVanish(GetFlag(sprm.Operand));
                    break;
                case 0x3:
                    newCHP.SetFcPic(sprm.Operand);
                    newCHP.SetFSpec(true);
                    break;
                case 0x4:
                    newCHP.SetIbstRMark((short)sprm.Operand);
                    break;
                case 0x5:
                    newCHP.SetDttmRMark(new DateAndTime(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x6:
                    newCHP.SetFData(GetFlag(sprm.Operand));
                    break;
                case 0x7:
                    //don't care about this
                    break;
                case 0x8:
                    //short chsDiff = (short)((param & 0xff0000) >>> 16);
                    int operand = sprm.Operand;
                    short chsDiff = (short)(operand & 0x0000ff);
                    newCHP.SetFChsDiff(GetFlag(chsDiff));
                    newCHP.SetChse((short)(operand & 0xffff00));
                    break;
                case 0x9:
                    newCHP.SetFSpec(true);
                    newCHP.SetFtcSym(LittleEndian.GetShort(sprm.Grpprl, sprm.GrpprlOffset));
                    newCHP.SetXchSym(LittleEndian.GetShort(sprm.Grpprl, sprm.GrpprlOffset + 2));
                    break;
                case 0xa:
                    newCHP.SetFOle2(GetFlag(sprm.Operand));
                    break;
                case 0xb:

                    // Obsolete
                    break;
                case 0xc:
                    newCHP.SetIcoHighlight((byte)sprm.Operand);
                    newCHP.SetFHighlight(GetFlag(sprm.Operand));
                    break;
                case 0xd:

                    //	undocumented
                    break;
                case 0xe:
                    newCHP.SetFcObj(sprm.Operand);
                    break;
                case 0xf:

                    // undocumented
                    break;
                case 0x10:

                    // undocumented
                    break;

                // undocumented till 0x30

                case 0x11:
                    // sprmCFWebHidden
                    break;
                case 0x12:
                    break;
                case 0x13:
                    break;
                case 0x14:
                    break;
                case 0x15:
                    // sprmCRsidProp
                    break;
                case 0x16:
                    // sprmCRsidText
                    break;
                case 0x17:
                    // sprmCRsidRMDel
                    break;
                case 0x18:
                    // sprmCFSpecVanish
                    break;
                case 0x19:
                    break;
                case 0x1a:
                    // sprmCFMathPr
                    break;
                case 0x1b:
                    break;
                case 0x1c:
                    break;
                case 0x1d:
                    break;
                case 0x1e:
                    break;
                case 0x1f:
                    break;
                case 0x20:
                    break;
                case 0x21:
                    break;
                case 0x22:
                    break;
                case 0x23:
                    break;
                case 0x24:
                    break;
                case 0x25:
                    break;
                case 0x26:
                    break;
                case 0x27:
                    break;
                case 0x28:
                    break;
                case 0x29:
                    break;
                case 0x2a:
                    break;
                case 0x2b:
                    break;
                case 0x2c:
                    break;
                case 0x2d:
                    break;
                case 0x2e:
                    break;
                case 0x2f:
                    break;
                case 0x30:
                    newCHP.SetIstd(sprm.Operand);
                    break;
                case 0x31:

                    //permutation vector for fast saves, who cares!
                    break;
                case 0x32:
                    newCHP.SetFBold(false);
                    newCHP.SetFItalic(false);
                    newCHP.SetFOutline(false);
                    newCHP.SetFStrike(false);
                    newCHP.SetFShadow(false);
                    newCHP.SetFSmallCaps(false);
                    newCHP.SetFCaps(false);
                    newCHP.SetFVanish(false);
                    newCHP.SetKul((byte)0);
                    newCHP.SetIco((byte)0);
                    break;
                case 0x33:
                    // preserve the fSpec Setting from the original CHP
                    bool fSpec = newCHP.IsFSpec();
                    newCHP = (CharacterProperties)oldCHP.Clone();
                    newCHP.SetFSpec(fSpec);

                    return;
                case 0x34:
                    // sprmCKcd
                    break;
                case 0x35:
                    newCHP.SetFBold(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFBold()));
                    break;
                case 0x36:
                    newCHP.SetFItalic(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFItalic()));
                    break;
                case 0x37:
                    newCHP.SetFStrike(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFStrike()));
                    break;
                case 0x38:
                    newCHP.SetFOutline(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFOutline()));
                    break;
                case 0x39:
                    newCHP.SetFShadow(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFShadow()));
                    break;
                case 0x3a:
                    newCHP.SetFSmallCaps(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFSmallCaps()));
                    break;
                case 0x3b:
                    newCHP.SetFCaps(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFCaps()));
                    break;
                case 0x3c:
                    newCHP.SetFVanish(GetCHPFlag((byte)sprm.Operand, oldCHP.IsFVanish()));
                    break;
                case 0x3d:
                    newCHP.SetFtcAscii((short)sprm.Operand);
                    break;
                case 0x3e:
                    newCHP.SetKul((byte)sprm.Operand);
                    break;
                case 0x3f:
                    operand = sprm.Operand;
                    int hps = operand & 0xff;
                    if (hps != 0)
                    {
                        newCHP.SetHps(hps);
                    }

                    //byte cInc = (byte)(((byte)(param & 0xfe00) >>> 4) >> 1);
                    byte cInc = (byte)((operand & 0xff00) >> 8);
                    cInc = (byte)(cInc >> 1);
                    if (cInc != 0)
                    {
                        newCHP.SetHps(Math.Max(newCHP.GetHps() + (cInc * 2), 2));
                    }

                    //byte hpsPos = (byte)((param & 0xff0000) >>> 8);
                    byte hpsPos = (byte)((operand & 0xff0000) >> 16);
                    if (hpsPos != 0x80)
                    {
                        newCHP.SetHpsPos(hpsPos);
                    }
                    bool fAdjust = (operand & 0x0100) > 0;
                    if (fAdjust && hpsPos != 128 && hpsPos != 0 && oldCHP.GetHpsPos() == 0)
                    {
                        newCHP.SetHps(Math.Max(newCHP.GetHps() + (-2), 2));
                    }
                    if (fAdjust && hpsPos == 0 && oldCHP.GetHpsPos() != 0)
                    {
                        newCHP.SetHps(Math.Max(newCHP.GetHps() + 2, 2));
                    }
                    break;
                case 0x40:
                    newCHP.SetDxaSpace(sprm.Operand);
                    break;
                case 0x41:
                    newCHP.SetLidDefault((short)sprm.Operand);
                    break;
                case 0x42:
                    newCHP.SetIco((byte)sprm.Operand);
                    break;
                case 0x43:
                    newCHP.SetHps(sprm.Operand);
                    break;
                case 0x44:
                    byte hpsLvl = (byte)sprm.Operand;
                    newCHP.SetHps(Math.Max(newCHP.GetHps() + (hpsLvl * 2), 2));
                    break;
                case 0x45:
                    newCHP.SetHpsPos((short)sprm.Operand);
                    break;
                case 0x46:
                    if (sprm.Operand != 0)
                    {
                        if (oldCHP.GetHpsPos() == 0)
                        {
                            newCHP.SetHps(Math.Max(newCHP.GetHps() + (-2), 2));
                        }
                    }
                    else
                    {
                        if (oldCHP.GetHpsPos() != 0)
                        {
                            newCHP.SetHps(Math.Max(newCHP.GetHps() + 2, 2));
                        }
                    }
                    break;
                case 0x47:
                    /*CharacterProperties genCHP = new CharacterProperties ();
                    genCHP.SetFtcAscii (4);
                    genCHP = (CharacterProperties) unCompressProperty (varParam, genCHP,
                      styleSheet);
                    CharacterProperties styleCHP = styleSheet.GetStyleDescription (oldCHP.
                      GetBaseIstd ()).GetCHP ();
                    if (genCHP.IsFBold () == newCHP.IsFBold ())
                    {
                      newCHP.SetFBold (styleCHP.IsFBold ());
                    }
                    if (genCHP.IsFItalic () == newCHP.IsFItalic ())
                    {
                      newCHP.SetFItalic (styleCHP.IsFItalic ());
                    }
                    if (genCHP.IsFSmallCaps () == newCHP.IsFSmallCaps ())
                    {
                      newCHP.SetFSmallCaps (styleCHP.IsFSmallCaps ());
                    }
                    if (genCHP.IsFVanish () == newCHP.IsFVanish ())
                    {
                      newCHP.SetFVanish (styleCHP.IsFVanish ());
                    }
                    if (genCHP.IsFStrike () == newCHP.IsFStrike ())
                    {
                      newCHP.SetFStrike (styleCHP.IsFStrike ());
                    }
                    if (genCHP.IsFCaps () == newCHP.IsFCaps ())
                    {
                      newCHP.SetFCaps (styleCHP.IsFCaps ());
                    }
                    if (genCHP.GetFtcAscii () == newCHP.GetFtcAscii ())
                    {
                      newCHP.SetFtcAscii (styleCHP.GetFtcAscii ());
                    }
                    if (genCHP.GetFtcFE () == newCHP.GetFtcFE ())
                    {
                      newCHP.SetFtcFE (styleCHP.GetFtcFE ());
                    }
                    if (genCHP.GetFtcOther () == newCHP.GetFtcOther ())
                    {
                      newCHP.SetFtcOther (styleCHP.GetFtcOther ());
                    }
                    if (genCHP.GetHps () == newCHP.GetHps ())
                    {
                      newCHP.SetHps (styleCHP.GetHps ());
                    }
                    if (genCHP.GetHpsPos () == newCHP.GetHpsPos ())
                    {
                      newCHP.SetHpsPos (styleCHP.GetHpsPos ());
                    }
                    if (genCHP.GetKul () == newCHP.GetKul ())
                    {
                      newCHP.SetKul (styleCHP.GetKul ());
                    }
                    if (genCHP.GetDxaSpace () == newCHP.GetDxaSpace ())
                    {
                      newCHP.SetDxaSpace (styleCHP.GetDxaSpace ());
                    }
                    if (genCHP.GetIco () == newCHP.GetIco ())
                    {
                      newCHP.SetIco (styleCHP.GetIco ());
                    }
                    if (genCHP.GetLidDefault () == newCHP.GetLidDefault ())
                    {
                      newCHP.SetLidDefault (styleCHP.GetLidDefault ());
                    }
                    if (genCHP.GetLidFE () == newCHP.GetLidFE ())
                    {
                      newCHP.SetLidFE (styleCHP.GetLidFE ());
                    }*/
                    break;
                case 0x48:
                    newCHP.SetIss((byte)sprm.Operand);
                    break;
                case 0x49:
                    newCHP.SetHps(LittleEndian.GetShort(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x4a:
                    int increment = LittleEndian.GetShort(sprm.Grpprl, sprm.GrpprlOffset);
                    newCHP.SetHps(Math.Max(newCHP.GetHps() + increment, 8));
                    break;
                case 0x4b:
                    newCHP.SetHpsKern(sprm.Operand);
                    break;
                case 0x4c:
                    //        unCompressCHPOperation (oldCHP, newCHP, 0x47, param, varParam,
                    //                                styleSheet, opSize);
                    break;
                case 0x4d:
                    float percentage = sprm.Operand / 100.0f;
                    int add = (int)(percentage * newCHP.GetHps());
                    newCHP.SetHps(newCHP.GetHps() + add);
                    break;
                case 0x4e:
                    newCHP.SetYsr((byte)sprm.Operand);
                    break;
                case 0x4f:
                    newCHP.SetFtcAscii((short)sprm.Operand);
                    break;
                case 0x50:
                    newCHP.SetFtcFE((short)sprm.Operand);
                    break;
                case 0x51:
                    newCHP.SetFtcOther((short)sprm.Operand);
                    break;
                case 0x52:
                    // sprmCCharScale
                    break;
                case 0x53:
                    newCHP.SetFDStrike(GetFlag(sprm.Operand));
                    break;
                case 0x54:
                    newCHP.SetFImprint(GetFlag(sprm.Operand));
                    break;
                case 0x55:
                    newCHP.SetFSpec(GetFlag(sprm.Operand));
                    break;
                case 0x56:
                    newCHP.SetFObj(GetFlag(sprm.Operand));
                    break;
                case 0x57:
                    byte[] buf = sprm.Grpprl;
                    int offset = sprm.GrpprlOffset;
                    newCHP.SetFPropMark(buf[offset]);
                    newCHP.SetIbstPropRMark(LittleEndian.GetShort(buf, offset + 1));
                    newCHP.SetDttmPropRMark(new DateAndTime(buf, offset + 3));
                    break;
                case 0x58:
                    newCHP.SetFEmboss(GetFlag(sprm.Operand));
                    break;
                case 0x59:
                    newCHP.SetSfxtText((byte)sprm.Operand);
                    break;
                case 0x5a:
                    // sprmCFBiDi
                    break;
                case 0x5b:
                    break;
                case 0x5c:
                    // sprmCFBoldBi
                    break;
                case 0x5d:
                    // sprmCFItalicBi
                    break;
                case 0x5e:
                    // sprmCFtcBi
                    break;
                case 0x5f:
                    // sprmCLidBi 
                    break;
                case 0x60:
                    // sprmCIcoBi
                    break;
                case 0x61:
                    // sprmCHpsBi
                    break;
                case 0x62:
                    byte[] xstDispFldRMark = new byte[32];
                    buf = sprm.Grpprl;
                    offset = sprm.GrpprlOffset;
                    newCHP.SetFDispFldRMark(buf[offset]);
                    newCHP.SetIbstDispFldRMark(LittleEndian.GetShort(buf, offset + 1));
                    newCHP.SetDttmDispFldRMark(new DateAndTime(buf, offset + 3));
                    Array.Copy(buf, offset + 7, xstDispFldRMark, 0, 32);
                    newCHP.SetXstDispFldRMark(xstDispFldRMark);
                    break;
                case 0x63:
                    newCHP.SetIbstRMarkDel((short)sprm.Operand);
                    break;
                case 0x64:
                    newCHP.SetDttmRMarkDel(new DateAndTime(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x65:
                    newCHP.SetBrc(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x66:
                    newCHP.SetShd(new ShadingDescriptor(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x67:
                    // Obsolete
                    break;
                case 0x68:
                    //  sprmCFUsePgsuSettings
                    break;
                case 0x69:
                    break;
                case 0x6a:
                    break;
                case 0x6b:
                    break;
                case 0x6c:
                    break;
                case 0x6d:
                    newCHP.SetLidDefault((short)sprm.Operand);
                    break;
                case 0x6e:
                    newCHP.SetLidFE((short)sprm.Operand);
                    break;
                case 0x6f:
                    newCHP.SetIdctHint((byte)sprm.Operand);
                    break;
                case 0x70:
                    newCHP.SetIco24(sprm.Operand);
                    break;
                case 0x71:
                    // sprmCShd
                    break;
                case 0x72:
                    // sprmCBrc
                    break;
                case 0x73:
                    // sprmCRgLid0
                    break;
                case 0x74:
                    // sprmCRgLid1
                    break;
            }
        }

        /**
         * Converts an int into a bool. If the int is non-zero, it returns true.
         * Otherwise it returns false.
         *
         * @param x The int to Convert.
         *
         * @return A bool whose value depends on x.
         */
        public static bool GetFlag(int x)
        {
            return x != 0;
        }

        private static bool GetCHPFlag(byte x, bool oldVal)
        {
            /*
                 switch(x)
                 {
             case 0:
               return false;
             case 1:
               return true;
             case (byte)0x80:
               return oldVal;
             case (byte)0x81:
               return !oldVal;
             default:
               return false;
                 }
             */
            if (x == 0)
            {
                return false;
            }
            else if (x == 1)
            {
                return true;
            }
            else if ((x & 0x81) == 0x80)
            {
                return oldVal;
            }
            else if ((x & 0x81) == 0x81)
            {
                return !oldVal;
            }
            else
            {
                return false;
            }
        }

    }
}


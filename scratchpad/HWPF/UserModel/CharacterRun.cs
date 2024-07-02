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

using NPOI.HWPF.SPRM;
using NPOI.HWPF.Model;
using System;
namespace NPOI.HWPF.UserModel
{

    /**
     * This class represents a run of text that share common properties.
     *
     * @author Ryan Ackley
     */
    public class CharacterRun : Range
    {
        public const short SPRM_FRMARKDEL = (short)0x0800;
        public const short SPRM_FRMARK = 0x0801;
        public const short SPRM_FFLDVANISH = 0x0802;
        public const short SPRM_PICLOCATION = 0x6A03;
        public const short SPRM_IBSTRMARK = 0x4804;
        public const short SPRM_DTTMRMARK = 0x6805;
        public const short SPRM_FDATA = 0x0806;
        public const short SPRM_SYMBOL = 0x6A09;
        public const short SPRM_FOLE2 = 0x080A;
        public const short SPRM_HIGHLIGHT = 0x2A0C;
        public const short SPRM_OBJLOCATION = 0x680E;
        public const short SPRM_ISTD = 0x4A30;
        public const short SPRM_FBOLD = 0x0835;
        public const short SPRM_FITALIC = 0x0836;
        public const short SPRM_FSTRIKE = 0x0837;
        public const short SPRM_FOUTLINE = 0x0838;
        public const short SPRM_FSHADOW = 0x0839;
        public const short SPRM_FSMALLCAPS = 0x083A;
        public const short SPRM_FCAPS = 0x083B;
        public const short SPRM_FVANISH = 0x083C;
        public const short SPRM_KUL = 0x2A3E;
        public const short SPRM_DXASPACE = unchecked((short)0x8840);
        public const short SPRM_LID = 0x4A41;
        public const short SPRM_ICO = 0x2A42;
        public const short SPRM_HPS = 0x4A43;
        public const short SPRM_HPSPOS = 0x4845;
        public const short SPRM_ISS = 0x2A48;
        public const short SPRM_HPSKERN = 0x484B;
        public const short SPRM_YSRI = 0x484E;
        public const short SPRM_RGFTCASCII = 0x4A4F;
        public const short SPRM_RGFTCFAREAST = 0x4A50;
        public const short SPRM_RGFTCNOTFAREAST = 0x4A51;
        public const short SPRM_CHARSCALE = 0x4852;
        public const short SPRM_FDSTRIKE = 0x2A53;
        public const short SPRM_FIMPRINT = 0x0854;
        public const short SPRM_FSPEC = 0x0855;
        public const short SPRM_FOBJ = 0x0856;
        public const short SPRM_PROPRMARK = unchecked((short)0xCA57);
        public const short SPRM_FEMBOSS = 0x0858;
        public const short SPRM_SFXTEXT = 0x2859;
        public const short SPRM_DISPFLDRMARK = unchecked((short)0xCA62);
        public const short SPRM_IBSTRMARKDEL = 0x4863;
        public const short SPRM_DTTMRMARKDEL = 0x6864;
        public const short SPRM_BRC = 0x6865;
        public const short SPRM_SHD = 0x4866;
        public const short SPRM_IDSIRMARKDEL = 0x4867;
        public const short SPRM_CPG = 0x486B;
        public const short SPRM_NONFELID = 0x486D;
        public const short SPRM_FELID = 0x486E;
        public const short SPRM_IDCTHINT = 0x286F;

        SprmBuffer _chpx;
        CharacterProperties _props;

        /**
         *
         * @param chpx The chpx this object is based on.
         * @param ss The stylesheet for the document this run belongs to.
         * @param istd The style index if this Run's base style.
         * @param parent The parent range of this character run (usually a paragraph).
         */
        internal CharacterRun(CHPX chpx, StyleSheet ss, short istd, Range parent)
            : base(Math.Max(parent._start, chpx.Start), Math.Min(parent._end, chpx.End), parent)
        {

            _props = chpx.GetCharacterProperties(ss, istd);
            _chpx = chpx.GetSprmBuf();
        }

        /**
         * Here for Runtime type determination using a switch statement convenient.
         *
         * @return TYPE_CHARACTER
         */
        public int type()
        {
            return TYPE_CHARACTER;
        }

        public bool IsMarkedDeleted()
        {
            return _props.IsFRMarkDel();
        }

        public void MarkDeleted(bool mark)
        {
            _props.SetFRMarkDel(mark);

            byte newVal = (byte)(mark ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FRMARKDEL, newVal);

        }

        public bool IsBold()
        {
            return _props.IsFBold();
        }

        public void SetBold(bool bold)
        {
            _props.SetFBold(bold);

            byte newVal = (byte)(bold ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FBOLD, newVal);

        }

        public bool IsItalic()
        {
            return _props.IsFItalic();
        }

        public void SetItalic(bool italic)
        {
            _props.SetFItalic(italic);

            byte newVal = (byte)(italic ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FITALIC, newVal);

        }

        public bool IsOutlined()
        {
            return _props.IsFOutline();
        }

        public void SetOutline(bool outlined)
        {
            _props.SetFOutline(outlined);

            byte newVal = (byte)(outlined ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FOUTLINE, newVal);

        }

        public bool IsFldVanished()
        {
            return _props.IsFFldVanish();
        }

        public void SetFldVanish(bool fldVanish)
        {
            _props.SetFFldVanish(fldVanish);

            byte newVal = (byte)(fldVanish ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FFLDVANISH, newVal);

        }

        public bool IsSmallCaps()
        {
            return _props.IsFSmallCaps();
        }

        public void SetSmallCaps(bool smallCaps)
        {
            _props.SetFSmallCaps(smallCaps);

            byte newVal = (byte)(smallCaps ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FSMALLCAPS, newVal);

        }

        public bool IsCapitalized()
        {
            return _props.IsFCaps();
        }

        public void SetCapitalized(bool caps)
        {
            _props.SetFCaps(caps);

            byte newVal = (byte)(caps ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FCAPS, newVal);

        }

        public bool IsVanished()
        {
            return _props.IsFVanish();
        }

        public void SetVanished(bool vanish)
        {
            _props.SetFVanish(vanish);

            byte newVal = (byte)(vanish ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FVANISH, newVal);

        }

        public bool IsMarkedInserted()
        {
            return _props.IsFRMark();
        }

        public void MarkInserted(bool mark)
        {
            _props.SetFRMark(mark);

            byte newVal = (byte)(mark ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FRMARK, newVal);

        }

        public bool IsStrikeThrough()
        {
            return _props.IsFStrike();
        }

        public void StrikeThrough(bool strike)
        {
            _props.SetFStrike(strike);

            byte newVal = (byte)(strike ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FSTRIKE, newVal);

        }

        public bool IsShadowed()
        {
            return _props.IsFShadow();
        }

        public void SetShadow(bool shadow)
        {
            _props.SetFShadow(shadow);

            byte newVal = (byte)(shadow ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FSHADOW, newVal);

        }

        public bool IsEmbossed()
        {
            return _props.IsFEmboss();
        }

        public void SetEmbossed(bool emboss)
        {
            _props.SetFEmboss(emboss);

            byte newVal = (byte)(emboss ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FEMBOSS, newVal);

        }

        public bool IsImprinted()
        {
            return _props.IsFImprint();
        }

        public void SetImprinted(bool imprint)
        {
            _props.SetFImprint(imprint);

            byte newVal = (byte)(imprint ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FIMPRINT, newVal);

        }

        public bool IsDoubleStrikeThrough()
        {
            return _props.IsFDStrike();
        }

        public void SetDoubleStrikethrough(bool dstrike)
        {
            _props.SetFDStrike(dstrike);

            byte newVal = (byte)(dstrike ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FDSTRIKE, newVal);

        }

        public void SetFtcAscii(int ftcAscii)
        {
            _props.SetFtcAscii(ftcAscii);

            _chpx.UpdateSprm(SPRM_RGFTCASCII, (short)ftcAscii);

        }

        public void SetFtcFE(int ftcFE)
        {
            _props.SetFtcFE(ftcFE);

            _chpx.UpdateSprm(SPRM_RGFTCFAREAST, (short)ftcFE);

        }

        public void SetFtcOther(int ftcOther)
        {
            _props.SetFtcOther(ftcOther);

            _chpx.UpdateSprm(SPRM_RGFTCNOTFAREAST, (short)ftcOther);

        }

        public int GetFontSize()
        {
            return _props.GetHps();
        }

        public void SetFontSize(int halfPoints)
        {
            _props.SetHps(halfPoints);

            _chpx.UpdateSprm(SPRM_HPS, (short)halfPoints);

        }

        public int GetCharacterSpacing()
        {
            return _props.GetDxaSpace();
        }

        public void SetCharacterSpacing(int twips)
        {
            _props.SetDxaSpace(twips);

            _chpx.UpdateSprm(SPRM_DXASPACE, twips);

        }

        public short GetSubSuperScriptIndex()
        {
            return _props.GetIss();
        }

        public void SetSubSuperScriptIndex(short iss)
        {
            _props.SetDxaSpace(iss);

            _chpx.UpdateSprm(SPRM_DXASPACE, iss);

        }

        public int GetUnderlineCode()
        {
            return _props.GetKul();
        }

        public void SetUnderlineCode(int kul)
        {
            _props.SetKul((byte)kul);
            _chpx.UpdateSprm(SPRM_KUL, (byte)kul);
        }

        public int GetColor()
        {
            return _props.GetIco();
        }

        public void SetColor(int color)
        {
            _props.SetIco((byte)color);
            _chpx.UpdateSprm(SPRM_ICO, (byte)color);
        }

        public int GetVerticalOffset()
        {
            return _props.GetHpsPos();
        }

        public void SetVerticalOffset(int hpsPos)
        {
            _props.SetHpsPos(hpsPos);
            _chpx.UpdateSprm(SPRM_HPSPOS, (byte)hpsPos);
        }

        public int GetKerning()
        {
            return _props.GetHpsKern();
        }

        public void SetKerning(int kern)
        {
            _props.SetHpsKern(kern);
            _chpx.UpdateSprm(SPRM_HPSKERN, (short)kern);
        }

        public bool IsHighlighted()
        {
            return _props.IsFHighlight();
        }

        public void SetHighlighted(byte color)
        {
            _props.SetFHighlight(true);
            _props.SetIcoHighlight(color);
            _chpx.UpdateSprm(SPRM_HIGHLIGHT, color);
        }

        public String GetFontName()
        {
            return _doc.GetFontTable().GetMainFont(_props.GetFtcAscii());
        }

        public bool IsSpecialCharacter()
        {
            return _props.IsFSpec();
        }

        public void SetSpecialCharacter(bool spec)
        {
            _props.SetFSpec(spec);

            byte newVal = (byte)(spec ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FSPEC, newVal);
        }

        public bool IsObj()
        {
            return _props.IsFObj();
        }

        public void SetObj(bool obj)
        {
            _props.SetFObj(obj);

            byte newVal = (byte)(obj ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FOBJ, newVal);
        }

        public int GetPicOffset()
        {
            return _props.GetFcPic();
        }

        public void SetPicOffset(int offset)
        {
            _props.SetFcPic(offset);
            _chpx.UpdateSprm(SPRM_PICLOCATION, offset);
        }

        /**
         * Does the picture offset represent picture
         *  or binary data?
         * If it's Set, then the picture offset refers to
         *  a NilPICFAndBinData structure, otherwise to a
         *  PICFAndOfficeArtData
         */
        public bool IsData()
        {
            return _props.IsFData();
        }

        public void SetData(bool data)
        {
            _props.SetFData(data);

            byte newVal = (byte)(data ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FOBJ, newVal);
        }

        public bool IsOle2()
        {
            return _props.IsFOle2();
        }

        public void SetOle2(bool ole)
        {
            _props.SetFOle2(ole);

            byte newVal = (byte)(ole ? 1 : 0);
            _chpx.UpdateSprm(SPRM_FOBJ, newVal);
        }

        public int GetObjOffset()
        {
            return _props.GetFcObj();
        }

        public void SetObjOffset(int obj)
        {
            _props.SetFcObj(obj);
            _chpx.UpdateSprm(SPRM_OBJLOCATION, obj);
        }

        /**
        * Get the ico24 field for the CHP record.
        */
        public int GetIco24()
        {
            return _props.GetIco24();
        }

        /**
         * Set the ico24 field for the CHP record.
         */
        public void SetIco24(int colour24)
        {
            _props.SetIco24(colour24);
        }

        /**
         * clone the CharacterProperties object associated with this
         * characterRun so that you can apply it to another CharacterRun
         */
        public CharacterProperties CloneProperties()
        {
            return (CharacterProperties)_props.Clone();

        }

        /**
         * Used to create a deep copy of this object.
         *
         * @return A deep copy.
         * @throws CloneNotSupportedException never
         */
        public override Object Clone()
        {
            CharacterRun cp = (CharacterRun)base.Clone();
            cp._props.SetDttmRMark((DateAndTime)_props.GetDttmRMark().Clone());
            cp._props.SetDttmRMarkDel((DateAndTime)_props.GetDttmRMarkDel().Clone());
            cp._props.SetDttmPropRMark((DateAndTime)_props.GetDttmPropRMark().Clone());
            cp._props.SetDttmDispFldRMark((DateAndTime)_props.GetDttmDispFldRMark().
                                          Clone());
            cp._props.SetXstDispFldRMark((byte[])_props.GetXstDispFldRMark().Clone());
            cp._props.SetShd((ShadingDescriptor)_props.GetShd().Clone());

            return cp;
        }

        /**
         * Returns true, if the CharacterRun is a special character run Containing a symbol, otherwise false.
         *
         * <p>In case of a symbol, the {@link #text()} method always returns a single character 0x0028, but word actually stores
         * the character in a different field. Use {@link #GetSymbolCharacter()} to get that character and {@link #GetSymbolFont()}
         * to determine its font.
         */
        public bool IsSymbol()
        {
            return IsSpecialCharacter() && Text.Equals("\u0028");
        }

        /**
         * Returns the symbol character, if this is a symbol character Run.
         * 
         * @see #isSymbol()
         * @throws InvalidOperationException If this is not a symbol character Run: call {@link #isSymbol()} first.
         */
        public char GetSymbolCharacter()
        {
            if (IsSymbol())
            {
                return (char)_props.GetXchSym();
            }
            else
                throw new InvalidOperationException("Not a symbol CharacterRun");
        }

        /**
         * Returns the symbol font, if this is a symbol character Run. Might return null, if the font index is not found in the font table.
         * 
         * @see #isSymbol()
         * @throws InvalidOperationException If this is not a symbol character Run: call {@link #isSymbol()} first.
         */
        public Ffn GetSymbolFont()
        {
            if (IsSymbol())
            {
                if(_doc.GetFontTable() == null)
                {
                    return null;
                }
                Ffn[] fontNames = _doc.GetFontTable().GetFontNames();

                if (fontNames.Length <= _props.GetFtcSym())
                    return null;

                return fontNames[_props.GetFtcSym()];
            }
            else
                throw new InvalidOperationException("Not a symbol CharacterRun");
        }

        public BorderCode GetBorder()
        {
            return _props.GetBrc();
        }

        public bool isHighlighted()
        {
            return _props.IsFHighlight();
        }
        public byte GetHighlightedColor()
        {
            return _props.GetIcoHighlight();
        }
        public void setHighlighted(byte color)
        {
            _props.SetFHighlight(true);
            _props.SetIcoHighlight(color);
            _chpx.UpdateSprm(SPRM_HIGHLIGHT, color);
        }
        public int getLanguageCode()
        {
            return _props.GetLidDefault();
        }
    }
}
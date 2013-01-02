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

using NPOI.HWPF.Model;
using System;
using NPOI.HWPF.SPRM;
namespace NPOI.HWPF.UserModel
{

    public class Paragraph : Range
    {
        public const short SPRM_JC = 0x2403;
        public const short SPRM_FSIDEBYSIDE = 0x2404;
        public const short SPRM_FKEEP = 0x2405;
        public const short SPRM_FKEEPFOLLOW = 0x2406;
        public const short SPRM_FPAGEBREAKBEFORE = 0x2407;
        public const short SPRM_BRCL = 0x2408;
        public const short SPRM_BRCP = 0x2409;
        public const short SPRM_ILVL = 0x260A;
        public const short SPRM_ILFO = 0x460B;
        public const short SPRM_FNOLINENUMB = 0x240C;
        public const short SPRM_CHGTABSPAPX = unchecked((short)0xC60D);
        public const short SPRM_DXARIGHT = unchecked((short)0x840E);
        public const short SPRM_DXALEFT = unchecked((short)0x840F);
        public const short SPRM_DXALEFT1 = unchecked((short)0x8411);
        public const short SPRM_DYALINE = 0x6412;
        public const short SPRM_DYABEFORE = unchecked((short)0xA413);
        public const short SPRM_DYAAFTER = unchecked((short)0xA414);
        public const short SPRM_CHGTABS = unchecked((short)0xC615);
        public const short SPRM_FINTABLE = 0x2416;
        public const short SPRM_FTTP = 0x2417;
        public const short SPRM_DXAABS = unchecked((short)0x8418);
        public const short SPRM_DYAABS = unchecked((short)0x8419);
        public const short SPRM_DXAWIDTH = unchecked((short)0x841A);
        public const short SPRM_PC = 0x261B;
        public const short SPRM_WR = 0x2423;
        public const short SPRM_BRCTOP = 0x6424;
        public const short SPRM_BRCLEFT = 0x6425;
        public const short SPRM_BRCBOTTOM = 0x6426;
        public const short SPRM_BRCRIGHT = 0x6427;
        public const short SPRM_BRCBAR = 0x6629;
        public const short SPRM_FNOAUTOHYPH = 0x242A;
        public const short SPRM_WHEIGHTABS = 0x442B;
        public const short SPRM_DCS = 0x442C;
        public const short SPRM_SHD = 0x442D;
        public const short SPRM_DYAFROMTEXT = unchecked((short)0x842E);
        public const short SPRM_DXAFROMTEXT = unchecked((short)0x842F);
        public const short SPRM_FLOCKED = 0x2430;
        public const short SPRM_FWIDOWCONTROL = 0x2431;
        public const short SPRM_RULER = unchecked((short)0xC632);
        public const short SPRM_FKINSOKU = 0x2433;
        public const short SPRM_FWORDWRAP = 0x2434;
        public const short SPRM_FOVERFLOWPUNCT = 0x2435;
        public const short SPRM_FTOPLINEPUNCT = 0x2436;
        public const short SPRM_AUTOSPACEDE = 0x2437;
        public const short SPRM_AUTOSPACEDN = 0x2438;
        public const short SPRM_WALIGNFONT = 0x4439;
        public const short SPRM_FRAMETEXTFLOW = 0x443A;
        public const short SPRM_ANLD = unchecked((short)0xC63E);
        public const short SPRM_PROPRMARK = unchecked((short)0xC63F);
        public const short SPRM_OUTLVL = 0x2640;
        public const short SPRM_FBIDI = 0x2441;
        public const short SPRM_FNUMRMLNS = 0x2443;
        public const short SPRM_CRLF = 0x2444;
        public const short SPRM_NUMRM = unchecked((short)0xC645);
        public const short SPRM_USEPGSUSETTINGS = 0x2447;
        public const short SPRM_FADJUSTRIGHT = 0x2448;


        internal short _istd;
        protected ParagraphProperties _props;
        internal SprmBuffer _papx;

        internal Paragraph(int startIdx, int endIdx, Table parent)
            : base(startIdx, endIdx, parent)
        {
            InitAll();
            PAPX papx = (PAPX)_paragraphs[_parEnd - 1];
            _props = papx.GetParagraphProperties(_doc.GetStyleSheet());
            _papx = papx.GetSprmBuf();
            _istd = papx.GetIstd();
        }

        internal Paragraph(PAPX papx, Range parent)
            : base(Math.Max(parent._start, papx.Start), Math.Min(parent._end, papx.End), parent)
        {

            _props = papx.GetParagraphProperties(_doc.GetStyleSheet());
            _papx = papx.GetSprmBuf();
            _istd = papx.GetIstd();
        }

        internal Paragraph(PAPX papx, Range parent, int start)
            : base(Math.Max(parent._start, start), Math.Min(parent._end, papx.End), parent)
        {

            _props = papx.GetParagraphProperties(_doc.GetStyleSheet());
            _papx = papx.GetSprmBuf();
            _istd = papx.GetIstd();
        }

        public short GetStyleIndex()
        {
            return _istd;
        }

        public override int Type
        {
            get
            {
                return TYPE_PARAGRAPH;
            }
        }

        public bool IsInTable()
        {
            return _props.GetFInTable();
        }

        public bool IsTableRowEnd()
        {
            return _props.GetFTtp() || _props.GetFTtpEmbedded();
        }

        public int GetTableLevel()
        {
            return _props.GetItap();
        }

        public bool IsEmbeddedCellMark()
        {
            return _props.GetFInnerTableCell();
        }

        public int GetJustification()
        {
            return _props.GetJc();
        }

        public void SetJustification(byte jc)
        {
            _props.SetJc(jc);
            _papx.UpdateSprm(SPRM_JC, jc);
        }

        public bool KeepOnPage()
        {
            return _props.GetFKeep();
        }

        public void SetKeepOnPage(bool fKeep)
        {
            _props.SetFKeep(fKeep);
            _papx.UpdateSprm(SPRM_FKEEP, Convert.ToByte(fKeep));
        }

        public bool KeepWithNext()
        {
            return _props.GetFKeepFollow();
        }

        public void SetKeepWithNext(bool fKeepFollow)
        {
            _props.SetFKeepFollow(fKeepFollow);
            _papx.UpdateSprm(SPRM_FKEEPFOLLOW, Convert.ToByte(fKeepFollow));
        }

        public bool PageBreakBefore()
        {
            return _props.GetFPageBreakBefore();
        }

        public void SetPageBreakBefore(bool fPageBreak)
        {
            _props.SetFPageBreakBefore(fPageBreak);
            _papx.UpdateSprm(SPRM_FPAGEBREAKBEFORE, Convert.ToByte(fPageBreak));
        }

        public bool IsLineNotNumbered()
        {
            return _props.GetFNoLnn();
        }

        public void SetLineNotNumbered(bool fNoLnn)
        {
            _props.SetFNoLnn(fNoLnn);
            _papx.UpdateSprm(SPRM_FNOLINENUMB, Convert.ToByte(fNoLnn));
        }

        public bool IsSideBySide()
        {
            return _props.GetFSideBySide();
        }

        public void SetSideBySide(bool fSideBySide)
        {
            byte sideBySide = (byte)(fSideBySide ? 1 : 0);
            _props.SetFSideBySide(fSideBySide);
            _papx.UpdateSprm(SPRM_FSIDEBYSIDE, sideBySide);
        }

        public bool IsAutoHyphenated
        {
            get
            {
                return !_props.GetFNoAutoHyph();
            }
            set 
            {
                _props.SetFNoAutoHyph(!value);
                _papx.UpdateSprm(SPRM_FNOAUTOHYPH, Convert.ToByte(!value));
           
            }
        }

        public bool IsWidowControlled()
        {
            return _props.GetFWidowControl();
        }

        public void SetWidowControl(bool widowControl)
        {
            _props.SetFWidowControl(widowControl);
            _papx.UpdateSprm(SPRM_FWIDOWCONTROL, Convert.ToByte(widowControl));
        }

        public int GetIndentFromRight()
        {
            return _props.GetDxaRight();
        }

        public void SetIndentFromRight(int dxaRight)
        {
            _props.SetDxaRight(dxaRight);
            _papx.UpdateSprm(SPRM_DXARIGHT, (short)dxaRight);
        }

        public int GetIndentFromLeft()
        {
            return _props.GetDxaLeft();
        }

        public void SetIndentFromLeft(int dxaLeft)
        {
            _props.SetDxaLeft(dxaLeft);
            _papx.UpdateSprm(SPRM_DXALEFT, (short)dxaLeft);
        }

        public int GetFirstLineIndent()
        {
            return _props.GetDxaLeft1();
        }

        public void SetFirstLineIndent(int first)
        {
            _props.SetDxaLeft1(first);
            _papx.UpdateSprm(SPRM_DXALEFT1, (short)first);
        }

        public LineSpacingDescriptor GetLineSpacing()
        {
            return _props.GetLspd();
        }

        public void SetLineSpacing(LineSpacingDescriptor lspd)
        {
            _props.SetLspd(lspd);
            _papx.UpdateSprm(SPRM_DYALINE, lspd.ToInt());
        }

        public int GetSpacingBefore()
        {
            return _props.GetDyaBefore();
        }

        public void SetSpacingBefore(int before)
        {
            _props.SetDyaBefore(before);
            _papx.UpdateSprm(SPRM_DYABEFORE, (short)before);
        }

        public int GetSpacingAfter()
        {
            return _props.GetDyaAfter();
        }

        public void SetSpacingAfter(int after)
        {
            _props.SetDyaAfter(after);
            _papx.UpdateSprm(SPRM_DYAAFTER, (short)after);
        }

        public bool IsKinsoku()
        {
            return _props.GetFKinsoku();
        }

        public void SetKinsoku(bool kinsoku)
        {
            byte kin = (byte)(kinsoku ? 1 : 0);
            _props.SetFKinsoku(kinsoku);
            _papx.UpdateSprm(SPRM_FKINSOKU, kin);
        }

        public bool IsWordWrapped()
        {
            return _props.GetFWordWrap();
        }

        public void SetWordWrapped(bool wrap)
        {
            byte wordWrap = (byte)(wrap ? 1 : 0);
            _props.SetFWordWrap(wrap);
            _papx.UpdateSprm(SPRM_FWORDWRAP, wordWrap);
        }

        public int GetFontAlignment()
        {
            return _props.GetWAlignFont();
        }

        public void SetFontAlignment(int align)
        {
            _props.SetWAlignFont(align);
            _papx.UpdateSprm(SPRM_WALIGNFONT, (short)align);
        }

        public bool IsVertical()
        {
            return _props.IsFVertical();
        }

        public void SetVertical(bool vertical)
        {
            _props.SetFVertical(vertical);
            _papx.UpdateSprm(SPRM_FRAMETEXTFLOW, GetFrameTextFlow());
        }

        public bool IsBackward()
        {
            return _props.IsFBackward();
        }

        public void SetBackward(bool bward)
        {
            _props.SetFBackward(bward);
            _papx.UpdateSprm(SPRM_FRAMETEXTFLOW, GetFrameTextFlow());
        }

        public virtual BorderCode GetTopBorder()
        {
            return _props.GetBrcTop();
        }

        public void SetTopBorder(BorderCode top)
        {
            _props.SetBrcTop(top);
            _papx.UpdateSprm(SPRM_BRCTOP, top.ToInt());
        }

        public virtual BorderCode GetLeftBorder()
        {
            return _props.GetBrcLeft();
        }

        public void SetLeftBorder(BorderCode left)
        {
            _props.SetBrcLeft(left);
            _papx.UpdateSprm(SPRM_BRCLEFT, left.ToInt());
        }

        public virtual BorderCode GetBottomBorder()
        {
            return _props.GetBrcBottom();
        }

        public void SetBottomBorder(BorderCode bottom)
        {
            _props.SetBrcBottom(bottom);
            _papx.UpdateSprm(SPRM_BRCBOTTOM, bottom.ToInt());
        }

        public virtual BorderCode GetRightBorder()
        {
            return _props.GetBrcRight();
        }

        public void SetRightBorder(BorderCode right)
        {
            _props.SetBrcRight(right);
            _papx.UpdateSprm(SPRM_BRCRIGHT, right.ToInt());
        }

        public virtual BorderCode GetBarBorder()
        {
            return _props.GetBrcBar();
        }

        public void SetBarBorder(BorderCode bar)
        {
            _props.SetBrcBar(bar);
            _papx.UpdateSprm(SPRM_BRCBAR, bar.ToInt());
        }

        public ShadingDescriptor GetShading()
        {
            return _props.GetShd();
        }

        public void SetShading(ShadingDescriptor shd)
        {
            _props.SetShd(shd);
            _papx.UpdateSprm(SPRM_SHD, shd.ToShort());
        }

        public DropCapSpecifier GetDropCap()
        {
            return _props.GetDcs();
        }

        public void SetDropCap(DropCapSpecifier dcs)
        {
            _props.SetDcs(dcs);
            _papx.UpdateSprm(SPRM_DCS, dcs.ToShort());
        }

        /**
         * Returns the ilfo, an index to the document's hpllfo, which
         *  describes the automatic number formatting of the paragraph.
         * A value of zero means it isn't numbered.
         */
        public int GetIlfo()
        {
            return _props.GetIlfo();
        }

        /**
         * Returns the multi-level indent for the paragraph. Will be
         *  zero for non-list paragraphs, and the first level of any
         *  list. Subsequent levels in hold values 1-8.
         */
        public int GetIlvl()
        {
            return _props.GetIlvl();
        }

        /**
         * Returns the heading level (1-8), or 9 if the paragraph
         *  isn't in a heading style.
         */
        public int GetLvl()
        {
            return _props.GetLvl();
        }

        internal void SetTableRowEnd(TableProperties props)
        {
            SetTableRowEnd(true);
            byte[] grpprl = TableSprmCompressor.CompressTableProperty(props);
            _papx.Append(grpprl);
        }

        private void SetTableRowEnd(bool val)
        {
            _props.SetFTtp(val);
            _papx.UpdateSprm(SPRM_FTTP, Convert.ToByte(val));
        }

        /**
         * clone the ParagraphProperties object associated with this Paragraph so
         * that you can apply the same properties to another paragraph.
         *
         */
        public ParagraphProperties CloneProperties()
        {
            return (ParagraphProperties)_props.Clone();
        }

        public override Object Clone()
        {
            Paragraph p = (Paragraph)base.Clone();
            p._props = (ParagraphProperties)_props.Clone();
            //p._baseStyle = _baseStyle;
            p._papx = new SprmBuffer();
            return p;
        }

        private short GetFrameTextFlow()
        {
            short retVal = 0;
            if (_props.IsFVertical())
            {
                retVal |= 1;
            }
            if (_props.IsFBackward())
            {
                retVal |= 2;
            }
            if (_props.IsFRotateFont())
            {
                retVal |= 4;
            }
            return retVal;
        }

    }
}

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

using NPOI.HWPF.Model.Types;
using System;
using NPOI.Util;
namespace NPOI.HWPF.UserModel
{

    public class ParagraphProperties : PAPAbstractType
    {

        private bool jcLogical = false;

        public ParagraphProperties()
        {
            SetAnld(new byte[84]);
            SetPhe(new byte[12]);
        }

        public override Object Clone()
        {

            ParagraphProperties pp = (ParagraphProperties)base.Clone();

            byte[] anld=GetAnld();
            byte[] anldcopy=new byte[anld.Length];
            Array.Copy(anld,anldcopy,anld.Length);
            pp.SetAnld(anldcopy);
            pp.SetBrcTop((BorderCode)GetBrcTop().Clone());
            pp.SetBrcLeft((BorderCode)GetBrcLeft().Clone());
            pp.SetBrcBottom((BorderCode)GetBrcBottom().Clone());
            pp.SetBrcRight((BorderCode)GetBrcRight().Clone());
            pp.SetBrcBetween((BorderCode)GetBrcBetween().Clone());
            pp.SetBrcBar((BorderCode)GetBrcBar().Clone());
            pp.SetDcs((DropCapSpecifier)GetDcs().Clone());
            pp.SetLspd((LineSpacingDescriptor)GetLspd().Clone());
            pp.SetShd((ShadingDescriptor)GetShd().Clone());
            byte[] phe = GetPhe();
            byte[] phecopy = new byte[phe.Length];
            Array.Copy(phe, phecopy, phe.Length);
            pp.SetPhe(phecopy);
            return pp;
        }

        public BorderCode GetBarBorder()
        {
            return base.GetBrcBar();
        }

        public BorderCode GetBottomBorder()
        {
            return base.GetBrcBottom();
        }

        public DropCapSpecifier GetDropCap()
        {
            return base.GetDcs();
        }

        public int GetFirstLineIndent()
        {
            return base.GetDxaLeft1();
        }

        public int GetFontAlignment()
        {
            return base.GetWAlignFont();
        }

        public int GetIndentFromLeft()
        {
            return base.GetDxaLeft();
        }

        public int GetIndentFromRight()
        {
            return base.GetDxaRight();
        }

        public int GetJustification()
        {
            if (jcLogical)
            {
                if (!GetFBiDi())
                    return GetJc();

                switch (GetJc())
                {
                    case 0:
                        return 2;
                    case 2:
                        return 0;
                    default:
                        return GetJc();
                }
            }

            return GetJc();
        }

        public BorderCode GetLeftBorder()
        {
            return base.GetBrcLeft();
        }

        public LineSpacingDescriptor GetLineSpacing()
        {
            return base.GetLspd();
        }

        public BorderCode GetRightBorder()
        {
            return base.GetBrcRight();
        }

        public ShadingDescriptor GetShading()
        {
            return base.GetShd();
        }

        public int GetSpacingAfter()
        {
            return base.GetDyaAfter();
        }

        public int GetSpacingBefore()
        {
            return base.GetDyaBefore();
        }

        public BorderCode GetTopBorder()
        {
            return base.GetBrcTop();
        }

        public bool IsAutoHyphenated()
        {
            return !base.GetFNoAutoHyph();
        }

        public bool IsBackward()
        {
            return base.IsFBackward();
        }

        public bool IsKinsoku()
        {
            return base.GetFKinsoku();
        }

        public bool IsLineNotNumbered()
        {
            return base.GetFNoLnn();
        }

        public bool IsSideBySide()
        {
            return base.GetFSideBySide();
        }

        public bool IsVertical()
        {
            return base.IsFVertical();
        }

        public bool IsWidowControlled()
        {
            return base.GetFWidowControl();
        }

        public bool IsWordWrapped()
        {
            return base.GetFWordWrap();
        }

        public bool keepOnPage()
        {
            return base.GetFKeep();
        }

        public bool keepWithNext()
        {
            return base.GetFKeepFollow();
        }

        public bool pageBreakBefore()
        {
            return base.GetFPageBreakBefore();
        }

        public void SetAutoHyphenated(bool auto)
        {
            base.SetFNoAutoHyph(!auto);
        }

        public void SetBackward(bool bward)
        {
            base.SetFBackward(bward);
        }

        public void SetBarBorder(BorderCode bar)
        {
            base.SetBrcBar(bar);
        }

        public void SetBottomBorder(BorderCode bottom)
        {
            base.SetBrcBottom(bottom);
        }

        public void SetDropCap(DropCapSpecifier dcs)
        {
            base.SetDcs(dcs);
        }

        public void SetFirstLineIndent(int first)
        {
            base.SetDxaLeft1(first);
        }

        public void SetFontAlignment(int align)
        {
            base.SetWAlignFont(align);
        }

        public void SetIndentFromLeft(int dxaLeft)
        {
            base.SetDxaLeft(dxaLeft);
        }

        public void SetIndentFromRight(int dxaRight)
        {
            base.SetDxaRight(dxaRight);
        }

        public void SetJustification(byte jc)
        {
            base.SetJc(jc);
            this.jcLogical = false;
        }

        public void SetJustificationLogical(byte jc)
        {
            base.SetJc(jc);
            this.jcLogical = true;
        }

        public void SetKeepOnPage(bool fKeep)
        {
            base.SetFKeep(fKeep);
        }

        public void SetKeepWithNext(bool fKeepFollow)
        {
            base.SetFKeepFollow(fKeepFollow);
        }

        public void SetKinsoku(bool kinsoku)
        {
            base.SetFKinsoku(kinsoku);
        }

        public void SetLeftBorder(BorderCode left)
        {
            base.SetBrcLeft(left);
        }

        public void SetLineNotNumbered(bool fNoLnn)
        {
            base.SetFNoLnn(fNoLnn);
        }

        public void SetLineSpacing(LineSpacingDescriptor lspd)
        {
            base.SetLspd(lspd);
        }

        public void SetPageBreakBefore(bool fPageBreak)
        {
            base.SetFPageBreakBefore(fPageBreak);
        }

        public void SetRightBorder(BorderCode right)
        {
            base.SetBrcRight(right);
        }

        public void SetShading(ShadingDescriptor shd)
        {
            base.SetShd(shd);
        }

        public void SetSideBySide(bool fSideBySide)
        {
            base.SetFSideBySide(fSideBySide);
        }

        public void SetSpacingAfter(int after)
        {
            base.SetDyaAfter(after);
        }

        public void SetSpacingBefore(int before)
        {
            base.SetDyaBefore(before);
        }

        public void SetTopBorder(BorderCode top)
        {
            base.SetBrcTop(top);
        }

        public void SetVertical(bool vertical)
        {
            base.SetFVertical(vertical);
        }

        public void SetWidowControl(bool widowControl)
        {
            base.SetFWidowControl(widowControl);
        }

        public void SetWordWrapped(bool wrap)
        {
            base.SetFWordWrap(wrap);
        }

    }


}
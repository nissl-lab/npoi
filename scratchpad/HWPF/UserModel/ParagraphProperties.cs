
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
    using NPOI.HWPF.Model;
    using NPOI.HWPF.SPRM;

    public class ParagraphProperties : PAPAbstractType  //, ICloneable
    {

        public ParagraphProperties()
        {
            field_21_lspd = new LineSpacingDescriptor();
            field_24_phe = new byte[12];
            field_46_brcTop = new BorderCode();
            field_47_brcLeft = new BorderCode();
            field_48_brcBottom = new BorderCode();
            field_49_brcRight = new BorderCode();
            field_50_brcBetween = new BorderCode();
            field_51_brcBar = new BorderCode();
            field_60_anld = new byte[84];
            this.field_17_fWidowControl = 1;
            this.field_21_lspd.SetMultiLinespace((short)1);
            this.field_21_lspd.SetDyaLine((short)240);
            //this.field_12_ilvl = (byte)9;
            this.field_58_lvl = (byte)9;
            this.field_66_rgdxaTab = new int[0];
            this.field_67_rgtbd = new byte[0];
            this.field_63_dttmPropRMark = new DateAndTime();

        }

        public int GetJustification()
        {
            return base.GetJc();
        }

        public void SetJustification(byte jc)
        {
            base.SetJc(jc);
        }

        public bool KeepOnPage()
        {
            return base.GetFKeep() != 0;
        }

        public void SetKeepOnPage(bool fKeep)
        {
            base.SetFKeep((byte)(fKeep ? 1 : 0));
        }

        public bool KeepWithNext()
        {
            return base.GetFKeepFollow() != 0;
        }

        public void SetKeepWithNext(bool fKeepFollow)
        {
            base.SetFKeepFollow((byte)(fKeepFollow ? 1 : 0));
        }

        public bool PageBreakBefore()
        {
            return base.GetFPageBreakBefore() != 0;
        }

        public void SetPageBreakBefore(bool fPageBreak)
        {
            base.SetFPageBreakBefore((byte)(fPageBreak ? 1 : 0));
        }

        public bool IsLineNotNumbered
        {
            get
            {
                return base.GetFNoLnn() != 0;
            }
            set 
            {
                base.SetFNoLnn((byte)(value ? 1 : 0));
            }
        }

        public bool IsSideBySide
        {
            get
            {
                return base.GetFSideBySide() != 0;
            }
            set 
            {
                base.SetFSideBySide((byte)(value ? 1 : 0));
            }
        }


        public bool IsAutoHyphenated
        {
            get
            {
                return base.GetFNoAutoHyph() == 0;
            }
            set 
            {
                base.SetFNoAutoHyph((byte)(!value ? 1 : 0));
            }
        }

        public bool IsWidowControlled
        {
            get
            {
                return base.GetFWidowControl() != 0;
            }
            set 
            {
                base.SetFWidowControl((byte)(value ? 1 : 0));
            }
        }

        public int GetIndentFromRight()
        {
            return base.GetDxaRight();
        }

        public void SetIndentFromRight(int dxaRight)
        {
            base.SetDxaRight(dxaRight);
        }

        public int GetIndentFromLeft()
        {
            return base.GetDxaLeft();
        }

        public void SetIndentFromLeft(int dxaLeft)
        {
            base.SetDxaLeft(dxaLeft);
        }

        public int GetFirstLineIndent()
        {
            return base.GetDxaLeft1();
        }

        public void SetFirstLineIndent(int first)
        {
            base.SetDxaLeft1(first);
        }

        public LineSpacingDescriptor GetLineSpacing()
        {
            return base.GetLspd();
        }

        public void SetLineSpacing(LineSpacingDescriptor lspd)
        {
            base.SetLspd(lspd);
        }

        public int GetSpacingBefore()
        {
            return base.GetDyaBefore();
        }

        public void SetSpacingBefore(int before)
        {
            base.SetDyaBefore(before);
        }

        public int GetSpacingAfter()
        {
            return base.GetDyaAfter();
        }

        public void SetSpacingAfter(int after)
        {
            base.SetDyaAfter(after);
        }

        public bool IsKinsoku
        {
            get
            {
                return base.GetFKinsoku() != 0;
            }
            set 
            {
                base.SetFKinsoku((byte)(value ? 1 : 0));
            }
        }


        public bool IsWordWrapped
        {
            get
            {
                return base.GetFWordWrap() != 0;
            }
            set 
            {
                base.SetFWordWrap((byte)(value ? 1 : 0));
            }
        }

        public int FontAlignment
        {
            get
            {
                return base.GetWAlignFont();
            }
            set 
            {
                base.SetWAlignFont(value);
            }
        }

        public bool IsVertical
        {
            get
            {
                return base.IsFVertical();
            }
            set 
            {
                base.SetFVertical(value);
            }
        }


        public bool IsBackward
        {
            get
            {
                return base.IsFBackward();
            }
            set 
            {
                base.SetFBackward(value);
            }
        }

        public BorderCode TopBorder
        {
            get
            {
                return base.GetBrcTop();
            }
            set 
            {
                base.SetBrcTop(value);
            }
        }

        public BorderCode LeftBorder
        {
            get
            {
                return base.GetBrcLeft();
            }
            set 
            {
                base.SetBrcLeft(value);
            }
        }

        public BorderCode BottomBorder
        {
            get
            {
                return base.GetBrcBottom();
            }
            set 
            {
                base.SetBrcBottom(value);
            }
        }

        public BorderCode RightBorder
        {
            get
            {
                return base.GetBrcRight();
            }
            set 
            {
                base.SetBrcRight(value);
            }
        }

        public BorderCode BarBorder
        {
            get
            {
                return base.GetBrcBar();
            }
            set 
            {
                base.SetBrcBar(value);
            }
        }

        public ShadingDescriptor GetShading()
        {
            return base.GetShd();
        }

        public void SetShading(ShadingDescriptor shd)
        {
            base.SetShd(shd);
        }

        public DropCapSpecifier GetDropCap()
        {
            return base.GetDcs();
        }

        public void SetDropCap(DropCapSpecifier dcs)
        {
            base.SetDcs(dcs);
        }

        //public Object Clone()
        //{
        //    ParagraphProperties pp = (ParagraphProperties)base.Clone();
        //    pp.field_21_lspd = (LineSpacingDescriptor)field_21_lspd.Clone();
        //    pp.field_24_phe = (byte[])field_24_phe.Clone();
        //    pp.field_46_brcTop = (BorderCode)field_46_brcTop.Clone();
        //    pp.field_47_brcLeft = (BorderCode)field_47_brcLeft.Clone();
        //    pp.field_48_brcBottom = (BorderCode)field_48_brcBottom.Clone();
        //    pp.field_49_brcRight = (BorderCode)field_49_brcRight.Clone();
        //    pp.field_50_brcBetween = (BorderCode)field_50_brcBetween.Clone();
        //    pp.field_51_brcBar = (BorderCode)field_51_brcBar.Clone();
        //    pp.field_60_anld = (byte[])field_60_anld.Clone();
        //    return pp;
        //}

    }
}
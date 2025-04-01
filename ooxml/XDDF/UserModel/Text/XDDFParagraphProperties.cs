/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Text
{
    using NPOI.Util;
    using NPOI.XDDF.UserModel;

    using NPOI.OpenXmlFormats.Dml;
    using System.Linq;

    public class XDDFParagraphProperties
    {
        private CT_TextParagraphProperties props;
        private XDDFParagraphBulletProperties bullet;
        internal XDDFParagraphProperties(CT_TextParagraphProperties properties)
        {
            this.props = properties;
            this.bullet = new XDDFParagraphBulletProperties(properties);
        }
        internal CT_TextParagraphProperties GetXmlObject()
        {
            return props;
        }

        public XDDFParagraphBulletProperties GetBulletProperties()
        {
            return bullet;
        }

        public int GetLevel()
        {
            if(props.lvlSpecified)
            {
                return 1 + props.lvl;
            }
            else
            {
                return 0;
            }
        }

        public void SetLevel(int? level)
        {
            if(level == null)
            {
                props.lvlSpecified = false;
            }
            else if(level < 1 || 9 < level)
            {
                throw new ArgumentException("Minimum inclusive: 1. Maximum inclusive: 9.");
            }
            else
            {
                props.lvl = level.Value - 1;
            }
        }

        /// <summary>
        /// </summary>
        /// <remarks>
        /// @since 4.0.1
        /// </remarks>
        public XDDFRunProperties AddDefaultRunProperties()
        {
            if(!props.IsSetDefRPr())
            {
                props.AddNewDefRPr();
            }
            return GetDefaultRunProperties();
        }

        public XDDFRunProperties GetDefaultRunProperties()
        {
            if(props.IsSetDefRPr())
            {
                return new XDDFRunProperties(props.defRPr);
            }
            else
            {
                return null;
            }
        }

        public void SetDefaultRunProperties(XDDFRunProperties properties)
        {
            if(properties == null)
            {
                if(props.IsSetDefRPr())
                {
                    props.UnsetDefRPr();
                }
            }
            else
            {
                props.defRPr = properties.GetXmlObject();
            }
        }

        public void SetEastAsianLineBreak(bool? value)
        {
            if(value == null)
            {
                props.eaLnBrkSpecified = false;
            }
            else
            {
                props.eaLnBrk = value.Value;
            }
        }

        public void SetLatinLineBreak(bool? value)
        {
            if(value == null)
            {
                props.latinLnBrkSpecified = false;
            }
            else
            {
                props.latinLnBrk = value.Value;
            }
        }

        public void SetHangingPunctuation(bool? value)
        {
            if(value == null)
            {
                props.hangingPunctSpecified = false;
            }
            else
            {
                props.hangingPunct = value.Value;
            }
        }

        public void SetRightToLeft(bool? value)
        {
            if(value == null)
            {
                props.rtlSpecified = false;
            }
            else
            {
                props.rtl = value.Value;
            }
        }

        public void SetFontAlignment(FontAlignment? align)
        {
            if(align == null)
            {
                if(props.IsSetFontAlgn())
                {
                    props.UnsetFontAlgn();
                }
            }
            else
            {
                props.fontAlgn = align.Value.ToST_TextFontAlignType();
            }
        }

        public void SetTextAlignment(TextAlignment? align)
        {
            if(align == null)
            {
                if(props.IsSetAlgn())
                {
                    props.UnsetAlgn();
                }
            }
            else
            {
                props.algn = align.Value.ToST_TextAlignType();
            }
        }

        public void SetDefaultTabSize(double? points)
        {
            if(points == null)
            {
                if(props.IsSetDefTabSz())
                {
                    props.UnsetDefTabSz();
                }
            }
            else
            {
                props.defTabSz = Units.ToEMU(points.Value);
            }
        }

        public void SetIndentation(double? points)
        {
            if(points == null)
            {
                if(props.IsSetIndent())
                {
                    props.UnsetIndent();
                }
            }
            else if(points < -4032 || 4032 < points)
            {
                throw new ArgumentException("Minimum inclusive = -4032. Maximum inclusive = 4032.");
            }
            else
            {
                props.indent = Units.ToEMU(points.Value);
            }
        }

        public void SetMarginLeft(double? points)
        {
            if(points == null)
            {
                if(props.IsSetMarL())
                {
                    props.UnsetMarL();
                }
            }
            else if(points < 0 || 4032 < points)
            {
                throw new ArgumentException("Minimum inclusive = 0. Maximum inclusive = 4032.");
            }
            else
            {
                props.marL = Units.ToEMU(points.Value);
            }
        }

        public void SetMarginRight(double? points)
        {
            if(points == null)
            {
                if(props.IsSetMarR())
                {
                    props.UnsetMarR();
                }
            }
            else if(points < 0 || 4032 < points)
            {
                throw new ArgumentException("Minimum inclusive = 0. Maximum inclusive = 4032.");
            }
            else
            {
                props.marR = Units.ToEMU(points.Value);
            }
        }

        public void SetLineSpacing(XDDFSpacing spacing)
        {
            if(spacing == null)
            {
                if(props.IsSetLnSpc())
                {
                    props.UnsetLnSpc();
                }
            }
            else
            {
                props.lnSpc = spacing.GetXmlObject();
            }
        }

        public void SetSpaceAfter(XDDFSpacing spacing)
        {
            if(spacing == null)
            {
                if(props.IsSetSpcAft())
                {
                    props.UnsetSpcAft();
                }
            }
            else
            {
                props.spcAft = spacing.GetXmlObject();
            }
        }

        public void SetSpaceBefore(XDDFSpacing spacing)
        {
            if(spacing == null)
            {
                if(props.IsSetSpcBef())
                {
                    props.UnsetSpcBef();
                }
            }
            else
            {
                props.spcBef = spacing.GetXmlObject();
            }
        }

        public XDDFTabStop AddTabStop()
        {
            if(!props.IsSetTabLst())
            {
                props.AddNewTabLst();
            }
            return new XDDFTabStop(props.tabLst.AddNewTab());
        }

        public XDDFTabStop InsertTabStop(int index)
        {
            if(!props.IsSetTabLst())
            {
                props.AddNewTabLst();
            }
            return new XDDFTabStop(props.tabLst.InsertNewTab(index));
        }

        public void RemoveTabStop(int index)
        {
            if(props.IsSetTabLst())
            {
                props.tabLst.RemoveTab(index);
            }
        }

        public XDDFTabStop GetTabStop(int index)
        {
            if(props.IsSetTabLst())
            {
                return new XDDFTabStop(props.tabLst.GetTabArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFTabStop> GetTabStops()
        {
            if(props.IsSetTabLst())
            {
                return [.. props.tabLst.tab.Select(p=> new XDDFTabStop(p))];
            }
            else
            {
                return new List<XDDFTabStop>();
            }
        }

        public int CountTabStops()
        {
            if(props.IsSetTabLst())
            {
                return props.tabLst.SizeOfTabArray();
            }
            else
            {
                return 0;
            }
        }

        public XDDFExtensionList ExtensionList
        {
            get
            {
                if(props.IsSetExtLst())
                {
                    return new XDDFExtensionList(props.extLst);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    if(props.IsSetExtLst())
                    {
                        props.UnsetExtLst();
                    }
                }
                else
                {
                    props.extLst = value.GetXmlObject();
                }
            }
        }
    }
}



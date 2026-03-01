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
    using NPOI.XDDF.UserModel;
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.Common.UserModel.Fonts;
    using System.Globalization;

    public class XDDFRunProperties
    {
        private CT_TextCharacterProperties props;

        public XDDFRunProperties()
                : this(new CT_TextCharacterProperties())
        {

        }
        internal XDDFRunProperties(CT_TextCharacterProperties properties)
        {
            this.props = properties;
        }
        internal CT_TextCharacterProperties GetXmlObject()
        {
            return props;
        }

        public void SetBaseline(int? value)
        {
            if(value == null)
            {
                props.baselineSpecified = false;
            }
            else
            {
                props.baseline = value.Value;
            }
        }

        public void SetDirty(bool? dirty)
        {
            if(dirty == null)
            {
                props.dirty = true;
            }
            else
            {
                props.dirty = dirty.Value;
            }
        }

        public void SetSpellError(bool? error)
        {
            if(error == null)
            {
                props.err = false;
            }
            else
            {
                props.err = error.Value;
            }
        }

        public void SetNoProof(bool? noproof)
        {
            if(noproof == null)
            {
                props.noProofSpecified = false;
            }
            else
            {
                props.noProof = noproof.Value;
            }
        }

        public void SetNormalizeHeights(bool? normalize)
        {
            if(normalize == null)
            {
                props.normalizeHSpecified = false;
            }
            else
            {
                props.normalizeH = normalize.Value;
            }
        }

        public void SetKumimoji(bool? kumimoji)
        {
            if(kumimoji == null)
            {
                props.kumimojiSpecified = false;
            }
            else
            {
                props.kumimoji = kumimoji.Value;
            }
        }

        public void SetBold(bool? bold)
        {
            if(bold == null)
            {
                props.bSpecified = false;
            }
            else
            {
                props.b = bold.Value;
            }
        }

        public void SetItalic(bool? italic)
        {
            if(italic == null)
            {
                props.iSpecified = false;
            }
            else
            {
                props.i = italic.Value;
            }
        }

        public void SetFontSize(double? size)
        {
            if(size == null)
            {
                props.szSpecified = false;
            }
            else if(size < 1 || 400 < size)
            {
                throw new ArgumentException("Minimum inclusive = 1. Maximum inclusive = 400.");
            }
            else
            {
                props.sz = (int) (100 * size);
            }
        }

        public void SetFillProperties(IXDDFFillProperties properties)
        {
            if(props.IsSetBlipFill())
            {
                props.UnsetBlipFill();
            }
            if(props.IsSetGradFill())
            {
                props.UnsetGradFill();
            }
            if(props.IsSetGrpFill())
            {
                props.UnsetGrpFill();
            }
            if(props.IsSetNoFill())
            {
                props.UnsetNoFill();
            }
            if(props.IsSetPattFill())
            {
                props.UnsetPattFill();
            }
            if(props.IsSetSolidFill())
            {
                props.UnsetSolidFill();
            }
            if(properties == null)
            {
                return;
            }
            if(properties is XDDFGradientFillProperties)
            {
                props.gradFill = ((XDDFGradientFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFGroupFillProperties)
            {
                props.grpFill = ((XDDFGroupFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFNoFillProperties)
            {
                props.noFill = ((XDDFNoFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFPatternFillProperties)
            {
                props.pattFill = ((XDDFPatternFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFPictureFillProperties)
            {
                props.blipFill = ((XDDFPictureFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFSolidFillProperties)
            {
                props.solidFill = ((XDDFSolidFillProperties) properties).GetXmlObject();
            }
        }

        public void SetCharacterKerning(double? kerning)
        {
            if(kerning == null)
            {
                props.kernSpecified = false;
            }
            else if(kerning < 0 || 4000 < kerning)
            {
                throw new ArgumentException("Minimum inclusive = 0. Maximum inclusive = 4000.");
            }
            else
            {
                props.kern = (int) (100 * kerning);
            }
        }

        public void SetCharacterSpacing(double? spacing)
        {
            if(spacing == null)
            {
                if(props.IsSetSpc())
                {
                    props.UnsetSpc();
                }
            }
            else if(spacing < -4000 || 4000 < spacing)
            {
                throw new ArgumentException("Minimum inclusive = -4000. Maximum inclusive = 4000.");
            }
            else
            {
                props.spc = (int) (100 * spacing);
            }
        }

        public void SetFonts(XDDFFont[] fonts)
        {
            foreach(XDDFFont font in fonts)
            {
                CT_TextFont xml = font.GetXmlObject();
                switch(font.Group)
                {
                    case FontGroup.COMPLEX_SCRIPT:
                        if(xml == null)
                        {
                            if(props.IsSetCs())
                            {
                                props.UnsetCs();
                            }
                        }
                        else
                        {
                            props.cs = xml;
                        }
                        break;
                    case FontGroup.EAST_ASIAN:
                        if(xml == null)
                        {
                            if(props.IsSetEa())
                            {
                                props.UnsetEa();
                            }
                        }
                        else
                        {
                            props.ea = xml;
                        }
                        break;
                    case FontGroup.LATIN:
                        if(xml == null)
                        {
                            if(props.IsSetLatin())
                            {
                                props.UnsetLatin();
                            }
                        }
                        else
                        {
                            props.latin = xml;
                        }
                        break;
                    case FontGroup.SYMBOL:
                        if(xml == null)
                        {
                            if(props.IsSetSym())
                            {
                                props.UnsetSym();
                            }
                        }
                        else
                        {
                            props.sym = xml;
                        }
                        break;
                }
            }
        }

        public void SetUnderline(UnderlineType? underline)
        {
            if(underline == null)
            {
                props.uSpecified = false;
            }
            else
            {
                props.u = underline.Value.ToST_TextUnderlineType();
            }
        }

        public void SetStrikeThrough(StrikeType? strike)
        {
            if(strike == null)
            {
                props.strikeSpecified = false;
            }
            else
            {
                props.strike = strike.Value.ToST_TextStrikeType();
            }
        }

        public void SetCapitals(CapsType? caps)
        {
            if(caps == null)
            {
                props.capSpecified = false;
            }
            else
            {
                props.cap = caps.Value.ToST_TextCapsType();
            }
        }

        public void SetHyperlink(XDDFHyperlink link)
        {
            if(link == null)
            {
                if(props.IsSetHlinkClick())
                {
                    props.UnsetHlinkClick();
                }
            }
            else
            {
                props.hlinkClick = link.GetXmlObject();
            }
        }

        public void SetMouseOver(XDDFHyperlink link)
        {
            if(link == null)
            {
                if(props.IsSetHlinkMouseOver())
                {
                    props.UnsetHlinkMouseOver();
                }
            }
            else
            {
                props.hlinkMouseOver = link.GetXmlObject();
            }
        }

        public void SetLanguage(CultureInfo lang)
        {
            if(lang == null)
            {
                props.lang = null;
            }
            else
            {
                props.lang = lang.Name;
            }
        }

        public void SetAlternativeLanguage(CultureInfo lang)
        {
            if(lang == null)
            {
                props.altLang = null;
            }
            else
            {
                props.altLang = lang.Name;
            }
        }

        public void SetHighlight(XDDFColor color)
        {
            if(color == null)
            {
                if(props.IsSetHighlight())
                {
                    props.UnsetHighlight();
                }
            }
            else
            {
                props.highlight = color.ColorContainer;
            }
        }

        public void SetLineProperties(XDDFLineProperties properties)
        {
            if(properties == null)
            {
                if(props.IsSetLn())
                {
                    props.UnsetLn();
                }
            }
            else
            {
                props.ln = properties.GetXmlObject();
            }
        }

        public void SetBookmark(string bookmark)
        {
            if(bookmark == null)
            {
                props.bmk = null;
            }
            else
            {
                props.bmk = bookmark;
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

        public XDDFEffectContainer EffectContainer
        {
            get
            {
                if(props.IsSetEffectDag())
                {
                    return new XDDFEffectContainer(props.effectDag);
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
                    if(props.IsSetEffectDag())
                    {
                        props.UnsetEffectDag();
                    }
                }
                else
                {
                    props.effectDag = value.GetXmlObject();
                }
            }
        }

        public XDDFEffectList EffectList
        {
            get
            {
                if(props.IsSetEffectLst())
                {
                    return new XDDFEffectList(props.effectLst);
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
                    if(props.IsSetEffectLst())
                    {
                        props.UnsetEffectLst();
                    }
                }
                else
                {
                    props.effectLst = value.GetXmlObject();
                }
            }
        }
    }
}

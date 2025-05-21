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
    using NPOI.OpenXmlFormats.Dml;
    using System.Linq;
    using NPOI.Util.Optional;
    using System.Globalization;

    public class XDDFTextBody
    {
        private CT_TextBody _body;
        private ITextContainer _parent;

        public XDDFTextBody(ITextContainer parent)
                : this(parent, new CT_TextBody())
        {
            Initialize();
        }
        public XDDFTextBody(ITextContainer parent, CT_TextBody body)
        {
            this._parent = parent;
            this._body = body;
        }
        public CT_TextBody GetXmlObject()
        {
            return _body;
        }

        public ITextContainer GetParentShape()
        {
            return _parent;
        }

        public XDDFTextParagraph Initialize()
        {
            _body.AddNewLstStyle();
            _body.AddNewBodyPr();
            XDDFBodyProperties bp = BodyProperties;
            bp.Anchoring = AnchorType.Top;
            bp.RightToLeft = false;
            XDDFTextParagraph p = AddNewParagraph();
            p.TextAlignment = TextAlignment.Left;
            XDDFRunProperties end = p.AddAfterLastRunProperties();
            end.SetLanguage(CultureInfo.GetCultureInfo("en-US"));
            end.SetFontSize(11.0);
            return p;
        }

        public void SetText(string text)
        {
            if(_body.SizeOfPArray() > 0)
            {
                // remove all but first paragraph
                for(int i = _body.SizeOfPArray() - 1; i > 0; i--)
                {
                    _body.RemoveP(i);
                }
                GetParagraph(0).Text = text;
            }
            else
            {
                // as there were no paragraphs yet, initialize the text body
                Initialize().Text = text;
            }
        }

        public XDDFTextParagraph AddNewParagraph()
        {
            return new XDDFTextParagraph(_body.AddNewP(), this);
        }

        public XDDFTextParagraph InsertNewParagraph(int index)
        {
            return new XDDFTextParagraph(_body.InsertNewP(index), this);
        }

        public void RemoveParagraph(int index)
        {
            _body.RemoveP(index);
        }

        public XDDFTextParagraph GetParagraph(int index)
        {
            return new XDDFTextParagraph(_body.GetPArray(index), this);
        }

        public List<XDDFTextParagraph> GetParagraphs()
        {
            return [.. _body.p.Select(x => new XDDFTextParagraph(x, this))];
        }

        public XDDFBodyProperties BodyProperties
        {
            get
            {
                return new XDDFBodyProperties(_body.bodyPr);
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    _body.AddNewBodyPr();
                }
                else
                {
                    _body.bodyPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties DefaultProperties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetDefPPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.defPPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetDefPPr())
                        {
                            style.UnsetDefPPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.defPPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level1Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl1pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl1pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl1pPr())
                        {
                            style.UnsetLvl1pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl1pPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level2Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl2pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl2pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl2pPr())
                        {
                            style.UnsetLvl2pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl2pPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level3Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl3pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl3pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl3pPr())
                        {
                            style.UnsetLvl3pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl3pPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level4Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl4pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl4pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl4pPr())
                        {
                            style.UnsetLvl4pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl4pPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level5Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl5pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl5pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl5pPr())
                        {
                            style.UnsetLvl5pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl5pPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level6Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl6pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl6pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl6pPr())
                        {
                            style.UnsetLvl6pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl6pPr = properties.GetXmlObject();
                }
            }
        }


        public XDDFParagraphProperties Level7Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl7pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl7pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl7pPr())
                        {
                            style.UnsetLvl7pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl7pPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level8Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl8pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl8pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl8pPr())
                        {
                            style.UnsetLvl8pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl8pPr = properties.GetXmlObject();
                }
            }
        }

        public XDDFParagraphProperties Level9Properties
        {
            get
            {
                if(_body.IsSetLstStyle() && _body.lstStyle.IsSetLvl9pPr())
                {
                    return new XDDFParagraphProperties(_body.lstStyle.lvl9pPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                var properties = value;
                if(properties == null)
                {
                    if(_body.IsSetLstStyle())
                    {
                        CT_TextListStyle style = _body.lstStyle;
                        if(style.IsSetLvl9pPr())
                        {
                            style.UnsetLvl9pPr();
                        }
                    }
                }
                else
                {
                    CT_TextListStyle style = _body.IsSetLstStyle() ? _body.lstStyle : _body.AddNewLstStyle();
                    style.lvl9pPr = properties.GetXmlObject();
                }
            }
        }

        internal Option<R> FindDefinedParagraphProperty<R>(Func<CT_TextParagraphProperties, Boolean> isSet,
            Func<CT_TextParagraphProperties, R> getter, int level) where R : class
        {
            if(_body.IsSetLstStyle() && level >= 0)
            {
                CT_TextListStyle list = _body.lstStyle;
                CT_TextParagraphProperties props = level == 0 ? list.defPPr : RetrieveProperties(list, level);
                if(props != null && isSet.Invoke(props))
                {
                    return Option<R>.Some(getter.Invoke(props));
                }
                else
                {
                    return FindDefinedParagraphProperty(isSet, getter, level - 1);
                }
            }
            else if(_parent != null)
            {
                return _parent.FindDefinedParagraphProperty(isSet, getter);
            }
            else
            {
                return Option<R>.None();
            }
        }
        internal Option<R> FindDefinedRunProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter, int level) where R : class
        {
            if(_body.IsSetLstStyle() && level >= 0)
            {
                CT_TextListStyle list = _body.lstStyle;
                CT_TextParagraphProperties props = level == 0 ? list.defPPr : RetrieveProperties(list, level);
                if(props != null && props.IsSetDefRPr() && isSet.Invoke(props.defRPr))
                {
                    return Option<R>.Some(getter.Invoke(props.defRPr));
                }
                else
                {
                    return FindDefinedRunProperty(isSet, getter, level - 1);
                }
            }
            else if(_parent != null)
            {
                return _parent.FindDefinedRunProperty(isSet, getter);
            }
            else
            {
                return Option<R>.None();
            }
        }


        internal ValueOption<R> FindDefinedParagraphValueProperty<R>(Func<CT_TextParagraphProperties, Boolean> isSet,
            Func<CT_TextParagraphProperties, R> getter, int level) where R : struct
        {
            if(_body.IsSetLstStyle() && level >= 0)
            {
                CT_TextListStyle list = _body.lstStyle;
                CT_TextParagraphProperties props = level == 0 ? list.defPPr : RetrieveProperties(list, level);
                if(props != null && isSet.Invoke(props))
                {
                    return ValueOption<R>.Some(getter.Invoke(props));
                }
                else
                {
                    return FindDefinedParagraphValueProperty(isSet, getter, level - 1);
                }
            }
            else if(_parent != null)
            {
                return _parent.FindDefinedParagraphValueProperty(isSet, getter);
            }
            else
            {
                return ValueOption<R>.None();
            }
        }
        internal ValueOption<R> FindDefinedRunValueProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter, int level) where R : struct
        {
            if(_body.IsSetLstStyle() && level >= 0)
            {
                CT_TextListStyle list = _body.lstStyle;
                CT_TextParagraphProperties props = level == 0 ? list.defPPr : RetrieveProperties(list, level);
                if(props != null && props.IsSetDefRPr() && isSet.Invoke(props.defRPr))
                {
                    return ValueOption<R>.Some(getter.Invoke(props.defRPr));
                }
                else
                {
                    return FindDefinedRunValueProperty(isSet, getter, level - 1);
                }
            }
            else if(_parent != null)
            {
                return _parent.FindDefinedRunValueProperty(isSet, getter);
            }
            else
            {
                return ValueOption<R>.None();
            }
        }

        private CT_TextParagraphProperties RetrieveProperties(CT_TextListStyle list, int level)
        {
            switch(level)
            {
                case 1:
                    if(list.IsSetLvl1pPr())
                    {
                        return list.lvl1pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 2:
                    if(list.IsSetLvl2pPr())
                    {
                        return list.lvl2pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 3:
                    if(list.IsSetLvl3pPr())
                    {
                        return list.lvl3pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 4:
                    if(list.IsSetLvl4pPr())
                    {
                        return list.lvl4pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 5:
                    if(list.IsSetLvl5pPr())
                    {
                        return list.lvl5pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 6:
                    if(list.IsSetLvl6pPr())
                    {
                        return list.lvl6pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 7:
                    if(list.IsSetLvl7pPr())
                    {
                        return list.lvl7pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 8:
                    if(list.IsSetLvl8pPr())
                    {
                        return list.lvl8pPr;
                    }
                    else
                    {
                        return null;
                    }
                case 9:
                    if(list.IsSetLvl9pPr())
                    {
                        return list.lvl9pPr;
                    }
                    else
                    {
                        return null;
                    }
                default:
                    return null;
            }
        }
    }
}

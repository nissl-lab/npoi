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
    using NPOI.Common.UserModel.Fonts;

    using NPOI.Util;


    using NPOI.XDDF.UserModel;

    using NPOI.OpenXmlFormats.Dml;
    public class XDDFParagraphBulletProperties
    {
        private CT_TextParagraphProperties props;
        internal XDDFParagraphBulletProperties(CT_TextParagraphProperties properties)
        {
            this.props = properties;
        }

        public IXDDFBulletStyle GetBulletStyle()
        {
            if(props.IsSetBuAutoNum())
            {
                return new XDDFBulletStyleAutoNumbered(props.buAutoNum);
            }
            else if(props.IsSetBuBlip())
            {
                return new XDDFBulletStylePicture(props.buBlip);
            }
            else if(props.IsSetBuChar())
            {
                return new XDDFBulletStyleCharacter(props.buChar);
            }
            else if(props.IsSetBuNone())
            {
                return new XDDFBulletStyleNone(props.buNone);
            }
            else
            {
                return null;
            }
        }

        public void SetBulletStyle(IXDDFBulletStyle style)
        {
            if(props.IsSetBuAutoNum())
            {
                props.UnsetBuAutoNum();
            }
            if(props.IsSetBuBlip())
            {
                props.UnsetBuBlip();
            }
            if(props.IsSetBuChar())
            {
                props.UnsetBuChar();
            }
            if(props.IsSetBuNone())
            {
                props.UnsetBuNone();
            }
            if(style != null)
            {
                if(style is XDDFBulletStyleAutoNumbered)
                {
                    props.buAutoNum = ((XDDFBulletStyleAutoNumbered) style).GetXmlObject();
                }
                else if(style is XDDFBulletStyleCharacter)
                {
                    props.buChar = ((XDDFBulletStyleCharacter) style).GetXmlObject();
                }
                else if(style is XDDFBulletStyleNone)
                {
                    props.buNone = ((XDDFBulletStyleNone) style).GetXmlObject();
                }
                else if(style is XDDFBulletStylePicture)
                {
                    props.buBlip = ((XDDFBulletStylePicture) style).GetXmlObject();
                }
            }
        }

        public XDDFColor BulletColor
        {
            get
            {
                if(props.IsSetBuClr())
                {
                    return XDDFColor.ForColorContainer(props.buClr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(props.IsSetBuClrTx())
                {
                    props.UnsetBuClrTx();
                }
                if(value == null)
                {
                    if(props.IsSetBuClr())
                    {
                        props.UnsetBuClr();
                    }
                }
                else
                {
                    props.buClr = value.ColorContainer;
                }
            }
        }

        public void SetBulletColorFollowText()
        {
            if(props.IsSetBuClr())
            {
                props.UnsetBuClr();
            }
            if(props.IsSetBuClrTx())
            {
                // nothing to do: already Set
            }
            else
            {
                props.AddNewBuClrTx();
            }
        }

        public XDDFFont GetBulletFont
        {
            get
            {
                if(props.IsSetBuFont())
                {
                    return new XDDFFont(FontGroup.SYMBOL, props.buFont);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(props.IsSetBuFontTx())
                {
                    props.UnsetBuFontTx();
                }
                if(value == null)
                {
                    if(props.IsSetBuFont())
                    {
                        props.UnsetBuFont();
                    }
                }
                else
                {
                    props.buFont = value.GetXmlObject();
                }
            }
        }


        public void SetBulletFontFollowText()
        {
            if(props.IsSetBuFont())
            {
                props.UnsetBuFont();
            }
            if(props.IsSetBuFontTx())
            {
                // nothing to do: already Set
            }
            else
            {
                props.AddNewBuFontTx();
            }
        }

        public IXDDFBulletSize BulletSize
        {
            get
            {
                if(props.IsSetBuSzPct())
                {
                    return new XDDFBulletSizePercent(props.buSzPct, null);
                }
                else if(props.IsSetBuSzPts())
                {
                    return new XDDFBulletSizePoints(props.buSzPts);
                }
                else if(props.IsSetBuSzTx())
                {
                    return new XDDFBulletSizeFollowText(props.buSzTx);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(props.IsSetBuSzPct())
                {
                    props.UnsetBuSzPct();
                }
                if(props.IsSetBuSzPts())
                {
                    props.UnsetBuSzPts();
                }
                if(props.IsSetBuSzTx())
                {
                    props.UnsetBuSzTx();
                }
                if(value != null)
                {
                    if(value is XDDFBulletSizeFollowText)
                    {
                        props.buSzTx = ((XDDFBulletSizeFollowText) value).GetXmlObject();
                    }
                    else if(value is XDDFBulletSizePercent)
                    {
                        props.buSzPct = ((XDDFBulletSizePercent) value).GetXmlObject();
                    }
                    else if(value is XDDFBulletSizePoints)
                    {
                        props.buSzPts = ((XDDFBulletSizePoints) value).GetXmlObject();
                    }
                }
            }
        }

        public void ClearAll()
        {
            if(props.IsSetBuAutoNum())
            {
                props.UnsetBuAutoNum();
            }
            if(props.IsSetBuBlip())
            {
                props.UnsetBuBlip();
            }
            if(props.IsSetBuChar())
            {
                props.UnsetBuChar();
            }
            if(props.IsSetBuNone())
            {
                props.UnsetBuNone();
            }
            if(props.IsSetBuClr())
            {
                props.UnsetBuClr();
            }
            if(props.IsSetBuClrTx())
            {
                props.UnsetBuClrTx();
            }
            if(props.IsSetBuFont())
            {
                props.UnsetBuFont();
            }
            if(props.IsSetBuFontTx())
            {
                props.UnsetBuFontTx();
            }
            if(props.IsSetBuSzPct())
            {
                props.UnsetBuSzPct();
            }
            if(props.IsSetBuSzPts())
            {
                props.UnsetBuSzPts();
            }
            if(props.IsSetBuSzTx())
            {
                props.UnsetBuSzTx();
            }
        }
    }
}

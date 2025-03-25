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


namespace NPOI.XDDF.UserModel
{

    using NPOI.OpenXmlFormats.Dml;


    public class XDDFSolidFillProperties : IXDDFFillProperties
    {
        private CT_SolidColorFillProperties props;

        public XDDFSolidFillProperties()
            : this(new CT_SolidColorFillProperties())
        {

        }

        public XDDFSolidFillProperties(XDDFColor color)
            : this(new CT_SolidColorFillProperties())
        {
            Color = color;
        }
        public XDDFSolidFillProperties(CT_SolidColorFillProperties properties)
        {
            this.props = properties;
        }
        public CT_SolidColorFillProperties GetXmlobject()
        {
            return props;
        }

        public XDDFColor Color
        {
            get
            {
                if(props.IsSetHslClr())
                {
                    return new XDDFColorHsl(props.hslClr);
                }
                else if(props.IsSetPrstClr())
                {
                    return new XDDFColorPreset(props.prstClr);
                }
                else if(props.IsSetSchemeClr())
                {
                    return new XDDFColorSchemeBased(props.schemeClr);
                }
                else if(props.IsSetScrgbClr())
                {
                    return new XDDFColorRgbPercent(props.scrgbClr);
                }
                else if(props.IsSetSrgbClr())
                {
                    return new XDDFColorRgbBinary(props.srgbClr);
                }
                else if(props.IsSetSysClr())
                {
                    return new XDDFColorSystemDefined(props.sysClr);
                }
                return null;
            }
            set
            {
                XDDFColor color = value;
                if(props.IsSetHslClr())
                {
                    props.UnsetHslClr();
                }
                if(props.IsSetPrstClr())
                {
                    props.UnsetPrstClr();
                }
                if(props.IsSetSchemeClr())
                {
                    props.UnsetSchemeClr();
                }
                if(props.IsSetScrgbClr())
                {
                    props.UnsetScrgbClr();
                }
                if(props.IsSetSrgbClr())
                {
                    props.UnsetSrgbClr();
                }
                if(props.IsSetSysClr())
                {
                    props.UnsetSysClr();
                }
                if(color == null)
                {
                    return;
                }
                if(color is XDDFColorHsl)
                {
                    props.hslClr = (CT_HslColor) color.GetXmlobject();
                }
                else if(color is XDDFColorPreset)
                {
                    props.prstClr = (CT_PresetColor) color.GetXmlobject();
                }
                else if(color is XDDFColorSchemeBased)
                {
                    props.schemeClr = (CT_SchemeColor) color.GetXmlobject();
                }
                else if(color is XDDFColorRgbPercent)
                {
                    props.scrgbClr = (CT_ScRgbColor) color.GetXmlobject();
                }
                else if(color is XDDFColorRgbBinary)
                {
                    props.srgbClr = (CT_SRgbColor) color.GetXmlobject();
                }
                else if(color is XDDFColorSystemDefined)
                {
                    props.sysClr = (CT_SystemColor) color.GetXmlobject();
                }
            }
        }
    }
}



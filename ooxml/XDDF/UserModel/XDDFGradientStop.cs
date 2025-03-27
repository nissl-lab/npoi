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
    public class XDDFGradientStop
    {
        private CT_GradientStop stop;
        public XDDFGradientStop(CT_GradientStop stop)
        {
            this.stop = stop;
        }
        public CT_GradientStop GetXmlObject()
        {
            return stop;
        }

        public int Position
        {
            get
            {
                return stop.pos;
            }
            set
            {
                stop.pos = value;
            }
        }

        public XDDFColor Color
        {
            get
            {
                if(stop.IsSetHslClr())
                {
                    return new XDDFColorHsl(stop.hslClr);
                }
                else if(stop.IsSetPrstClr())
                {
                    return new XDDFColorPreset(stop.prstClr);
                }
                else if(stop.IsSetSchemeClr())
                {
                    return new XDDFColorSchemeBased(stop.schemeClr);
                }
                else if(stop.IsSetScrgbClr())
                {
                    return new XDDFColorRgbPercent(stop.scrgbClr);
                }
                else if(stop.IsSetSrgbClr())
                {
                    return new XDDFColorRgbBinary(stop.srgbClr);
                }
                else if(stop.IsSetSysClr())
                {
                    return new XDDFColorSystemDefined(stop.sysClr);
                }
                return null;
            }
            set
            {
                XDDFColor color = value;
                if(stop.IsSetHslClr())
                {
                    stop.UnsetHslClr();
                }
                if(stop.IsSetPrstClr())
                {
                    stop.UnsetPrstClr();
                }
                if(stop.IsSetSchemeClr())
                {
                    stop.UnsetSchemeClr();
                }
                if(stop.IsSetScrgbClr())
                {
                    stop.UnsetScrgbClr();
                }
                if(stop.IsSetSrgbClr())
                {
                    stop.UnsetSrgbClr();
                }
                if(stop.IsSetSysClr())
                {
                    stop.UnsetSysClr();
                }
                if(color == null)
                {
                    return;
                }
                if(color is XDDFColorHsl)
                {
                    stop.hslClr = (CT_HslColor) color.GetXmlObject();
                }
                else if(color is XDDFColorPreset)
                {
                    stop.prstClr = (CT_PresetColor) color.GetXmlObject();
                }
                else if(color is XDDFColorSchemeBased)
                {
                    stop.schemeClr = (CT_SchemeColor) color.GetXmlObject();
                }
                else if(color is XDDFColorRgbPercent)
                {
                    stop.scrgbClr = (CT_ScRgbColor) color.GetXmlObject();
                }
                else if(color is XDDFColorRgbBinary)
                {
                    stop.srgbClr = (CT_SRgbColor) color.GetXmlObject();
                }
                else if(color is XDDFColorSystemDefined)
                {
                    stop.sysClr = (CT_SystemColor) color.GetXmlObject();
                }
            }
        }
    }
}



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
    public abstract class XDDFColor
    {
        protected CT_Color container;
        protected XDDFColor(CT_Color container)
        {
            this.container = container;
        }

        public static XDDFColor From(byte[] color)
        {
            return new XDDFColorRgbBinary(color);
        }

        public static XDDFColor From(int red, int green, int blue)
        {
            return new XDDFColorRgbPercent(red, green, blue);
        }

        public static XDDFColor From(PresetColor color)
        {
            return new XDDFColorPreset(color);
        }

        public static XDDFColor From(SchemeColor color)
        {
            return new XDDFColorSchemeBased(color);
        }

        public static XDDFColor From(SystemColor color)
        {
            return new XDDFColorSystemDefined(color);
        }
        public static XDDFColor ForColorContainer(CT_Color container)
        {
            if(container.hslClr != null)
            {
                return new XDDFColorHsl(container.hslClr, container);
            }
            else if(container.prstClr != null)
            {
                return new XDDFColorPreset(container.prstClr, container);
            }
            else if(container.schemeClr != null)
            {
                return new XDDFColorSchemeBased(container.schemeClr, container);
            }
            else if(container.scrgbClr != null)
            {
                return new XDDFColorRgbPercent(container.scrgbClr, container);
            }
            else if(container.IsSetSrgbClr())
            {
                return new XDDFColorRgbBinary(container.srgbClr, container);
            }
            else if(container.IsSetSysClr())
            {
                return new XDDFColorSystemDefined(container.sysClr, container);
            }
            return null;
        }
        public CT_Color ColorContainer
        {
            get
            {
                return container;
            }
            
        }
        public abstract object GetXmlobject();
    }
}

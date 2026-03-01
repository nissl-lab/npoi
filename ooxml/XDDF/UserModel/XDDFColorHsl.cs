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

    public class XDDFColorHsl : XDDFColor
    {
        private CT_HslColor color;

        public XDDFColorHsl(int hue, int saturation, int luminance)
            : this(new CT_HslColor(), new CT_Color())
        {
            Hue = hue;
            Saturation = saturation;
            Luminance = luminance;
        }
        public XDDFColorHsl(CT_HslColor color)
            : this(color, null)
        {

        }
        public XDDFColorHsl(CT_HslColor color, CT_Color container)
            : base(container)
        {

            this.color = color;
        }
        public override object GetXmlObject()
        {
            return color;
        }

        public int Hue
        {
            get
            {
                return color.hue;
            }
            set
            {
                color.hue = value;
            }
        }

        public int Saturation
        {
            get
            {
                return color.sat;
            }
            set
            {
                color.sat = value;
            }
        }

        public int Luminance
        {
            get
            {
                return color.lum;
            }
            set
            {
                color.lum = value;
            }
        }
    }
}

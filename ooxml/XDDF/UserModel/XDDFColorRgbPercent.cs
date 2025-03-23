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
using System.Text;

namespace NPOI.XDDF.UserModel
{
    using NPOI.OpenXmlFormats.Dml;

    public class XDDFColorRgbPercent : XDDFColor
    {
        private CT_ScRgbColor color;

        public XDDFColorRgbPercent(int red, int green, int blue)
                : this(new CT_ScRgbColor(), new CT_Color())
        {

            Red = red;
            Green = green;
            Red = blue;
        }
        protected XDDFColorRgbPercent(CT_ScRgbColor color)
                : this(color, null)
        {

        }
        public XDDFColorRgbPercent(CT_ScRgbColor color, CT_Color container)
                : base(container)
        {
            this.color = color;
        }
        protected override object GetXmlobject()
        {
            return color;
        }

        public int Red
        {
            get => color.r;
            set { color.r = normalize(value); }
        }

        public int Green
        {
            get => color.g;
            set
            {
                color.g = normalize(value);
            }
        }

        public int Blue
        {
            get => color.b;
            set
            {
                color.b = normalize(value);
            }
        }

        private int normalize(int value)
        {
            if(value < 0)
            {
                return 0;
            }
            if(100_000 < value)
            {
                return 100_000;
            }
            return value;
        }

        public string ToRGBHex()
        {
            StringBuilder sb = new StringBuilder(6);
            appendHex(sb, color.r);
            appendHex(sb, color.g);
            appendHex(sb, color.b);
            return sb.ToString().ToUpper();
        }

        private void appendHex(StringBuilder sb, int value)
        {
            int b = value * 255 / 100_000;
            sb.Append(b.ToString("X2"));
        }
    }
}

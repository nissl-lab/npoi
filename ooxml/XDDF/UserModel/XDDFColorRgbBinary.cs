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

    public class XDDFColorRgbBinary : XDDFColor
    {
        private CT_SRgbColor color;

        public XDDFColorRgbBinary(byte[] color)
            : this(new CT_SRgbColor(), new CT_Color())
        {
            Value = color;
        }
        public XDDFColorRgbBinary(CT_SRgbColor color)
            : this(color, null)
        {

        }
        public XDDFColorRgbBinary(CT_SRgbColor color, CT_Color container)
            : base(container)
        {
            this.color = color;
        }
        public override object GetXmlobject()
        {
            return color;
        }

        public byte[] Value
        {
            get { return color.val; }
            set { color.val = value; }
        }

        public string ToRGBHex()
        {
            StringBuilder sb = new StringBuilder(6);
            foreach(byte b in color.val)
            {
                sb.Append(b.ToString("X2"));
            }
            return sb.ToString().ToUpper();
        }
    }
}

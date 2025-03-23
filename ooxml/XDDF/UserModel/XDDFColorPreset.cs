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

    public class XDDFColorPreset : XDDFColor
    {
        private CT_PresetColor color;

        public XDDFColorPreset(PresetColor color)
            : this(new CT_PresetColor(), new CT_Color())
        {
            Value = color;
        }
        protected XDDFColorPreset(CT_PresetColor color)
                : this(color, null)
        {

        }
        public XDDFColorPreset(CT_PresetColor color, CT_Color container)
                : base(container)
        {

            this.color = color;
        }
        protected override object GetXmlobject()
        {
            return color;
        }

        public PresetColor? Value
        {
            get
            {
                if(color.valSpecified)
                {
                    return PresetColorExtensions.ValueOf(color.val);
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
                    color.valSpecified = false;
                }
                else
                {
                    color.val = value.Value.ToST_PresetColorVal();
                }
            }
        }
    }
}



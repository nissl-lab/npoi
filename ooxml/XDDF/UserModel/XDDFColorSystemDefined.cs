﻿/* ====================================================================
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

    public class XDDFColorSystemDefined : XDDFColor
    {
        private CT_SystemColor color;

        public XDDFColorSystemDefined(SystemColor color)
                : this(new CT_SystemColor(), new CT_Color())
        {

            Value = color;
        }
        public XDDFColorSystemDefined(CT_SystemColor color)
                : this(color, null)
        {

        }
        public XDDFColorSystemDefined(CT_SystemColor color, CT_Color container)
                : base(container)
        {

            this.color = color;
        }
        public override object GetXmlObject()
        {
            return color;
        }

        public SystemColor Value
        {
            get
            {
                return SystemColorExtensions.ValueOf(color.val);
            }
            set
            {
                color.val = value.ToST_SystemColorVal();
            }
        }

        public byte[] LastColor
        {
            get
            {
                if(color.lastClrSpecified)
                {
                    return color.lastClr;
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
                    if(color.lastClrSpecified)
                    {
                        color.lastClrSpecified = false;
                    }
                }
                else
                {
                    color.lastClr = value;
                }
            }
        }
    }
}



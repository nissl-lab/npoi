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
    public class XDDFLineEndProperties
    {
        private CT_LineEndProperties props;

        public XDDFLineEndProperties(CT_LineEndProperties properties)
        {
            this.props = properties;
        }
        public CT_LineEndProperties GetXmlObject()
        {
            return props;
        }

        public LineEndLength Length
        {
            get
            {
                return LineEndLengthExtensions.ValueOf(props.len);
            }
            set 
            {
                props.len = value.ToST_LineEndLength();
            }
        }

        public LineEndType Type
        {
            get
            {
                return LineEndTypeExtensions.ValueOf(props.type);
            }
            set
            {
                props.type = value.ToST_LineEndType();
            }
        }

        public LineEndWidth Width
        {
            get
            {
                return LineEndWidthExtensions.ValueOf(props.w);
            }
            set
            {
                props.w = value.ToST_LineEndWidth();
            }
        }
    }
}

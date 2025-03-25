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
    public class XDDFStretchInfoProperties
    {
        private CT_StretchInfoProperties props;

        public XDDFStretchInfoProperties(CT_StretchInfoProperties properties)
        {
            this.props = properties;
        }
        public CT_StretchInfoProperties GetXmlobject()
        {
            return props;
        }

        public XDDFRelativeRectangle GetFillRectangle()
        {
            if(props.IsSetFillRect())
            {
                return new XDDFRelativeRectangle(props.fillRect);
            }
            else
            {
                return null;
            }
        }

        public void SetFillRectangle(XDDFRelativeRectangle rectangle)
        {
            if(rectangle == null)
            {
                if(props.IsSetFillRect())
                {
                    props.UnsetFillRect();
                }
            }
            else
            {
                props.fillRect = rectangle.GetXmlobject();
            }
        }
    }
}



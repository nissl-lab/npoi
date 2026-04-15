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


using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XDDF.UserModel
{
    public class XDDFLinearShadeProperties
    {
        private CT_LinearShadeProperties props;

        public XDDFLinearShadeProperties(CT_LinearShadeProperties properties)
        {
            this.props = properties;
        }
        public CT_LinearShadeProperties GetXmlObject()
        {
            return props;
        }

        public double? Angle
        {
            get
            {
                if(props.angSpecified)
                {
                    return Angles.AttributeToDegrees(props.ang);
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
                    props.angSpecified = false;
                }
                else
                {
                    if(value < 0.0 || 360.0 <= value)
                    {
                        throw new System.ArgumentException("angle must be in the range [0, 360).");
                    }
                    props.ang = Angles.DegreesToAttribute(value.Value);
                }
            }
        }

        public bool? Scaled
        {
            get
            {
                if(props.scaledSpecified)
                {
                    return props.scaled;
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
                    props.scaledSpecified = false;
                }
                else
                {
                    props.scaled = value.Value;
                }
            }
        }
    }
}

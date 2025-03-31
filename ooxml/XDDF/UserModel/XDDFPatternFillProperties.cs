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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel
{
    using NPOI.Util;


    using NPOI.OpenXmlFormats.Dml;
    public class XDDFPatternFillProperties : IXDDFFillProperties
    {
        private CT_PatternFillProperties props;

        public XDDFPatternFillProperties()
                : this(new CT_PatternFillProperties())
        {
        }

        public XDDFPatternFillProperties(CT_PatternFillProperties properties)
        {
            this.props = properties;
        }
        public CT_PatternFillProperties GetXmlObject()
        {
            return props;
        }

        public PresetPattern? PresetPattern
        {
            get
            {
                if(props.prstSpecified)
                {
                    return PresetPatternExtensions.ValueOf(props.prst);
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
                    props.prstSpecified = false;   
                }
                else
                {
                    props.prst = value.Value.ToST_PresetPatternVal();
                }
            }
        }

        public XDDFColor BackgroundColor
        {
            get
            {
                if(props.IsSetBgClr())
                {
                    return XDDFColor.ForColorContainer(props.bgClr);
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
                    if(props.IsSetBgClr())
                    {
                        props.UnsetBgClr();
                    }
                }
                else
                {
                    props.bgClr = value.ColorContainer;
                }
            }
        }


        public XDDFColor ForegroundColor
        {
            get
            {
                if(props.IsSetFgClr())
                {
                    return XDDFColor.ForColorContainer(props.fgClr);
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
                    if(props.IsSetFgClr())
                    {
                        props.UnsetFgClr();
                    }
                }
                else
                {
                    props.fgClr = value.ColorContainer;
                }
            }
        }
    }
}

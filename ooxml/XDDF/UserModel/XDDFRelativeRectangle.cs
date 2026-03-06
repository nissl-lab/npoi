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
    public class XDDFRelativeRectangle
    {
        private CT_RelativeRect rect;

        public XDDFRelativeRectangle()
                : this(new CT_RelativeRect())
        {

        }

        public XDDFRelativeRectangle(CT_RelativeRect rectangle)
        {
            this.rect = rectangle;
        }
        public CT_RelativeRect GetXmlObject()
        {
            return rect;
        }

        public int? Bottom
        {
            get
            {
                if(rect.bSpecified)
                {
                    return rect.b;
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
                    if(rect.bSpecified)
                    {
                        rect.bSpecified = false;
                    }
                }
                else
                {
                    rect.b = value.Value;
                }
            }
        }

        public int? Left
        {
            get
            {
                if(rect.lSpecified)
                {
                    return rect.l;
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
                    if(rect.lSpecified)
                    {
                        rect.lSpecified = false;
                    }
                }
                else
                {
                    rect.l = value.Value;
                }
            }
        }

        public int? Right
        {
            get
            {
                if(rect.rSpecified)
                {
                    return rect.r;
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
                    if(rect.rSpecified)
                    {
                        rect.rSpecified = false;
                    }
                }
                else
                {
                    rect.r = value.Value;
                }
            }
        }

        public int? Top
        {
            get
            {
                if(rect.tSpecified)
                {
                    return rect.t;
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
                    if(rect.tSpecified)
                    {
                        rect.tSpecified = false;
                    }
                }
                else
                {
                    rect.t = value.Value;
                }
            }
        }
    }
}

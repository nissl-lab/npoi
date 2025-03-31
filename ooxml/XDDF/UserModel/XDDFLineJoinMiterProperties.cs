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
    public class XDDFLineJoinMiterProperties : IXDDFLineJoinProperties
    {
        private CT_LineJoinMiterProperties join;

        public XDDFLineJoinMiterProperties()
            : this(new CT_LineJoinMiterProperties())
        {

        }

        public XDDFLineJoinMiterProperties(CT_LineJoinMiterProperties join)
        {
            this.join = join;
        }
        public CT_LineJoinMiterProperties GetXmlObject()
        {
            return join;
        }

        public int? Limit
        {
            get
            {
                if(join.limSpecified)
                {
                    return join.lim;
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
                    if(join.limSpecified)
                    {
                        join.limSpecified = false;
                    }
                }
                else
                {
                    join.lim = value.Value;
                }
            }
        }
    }
}



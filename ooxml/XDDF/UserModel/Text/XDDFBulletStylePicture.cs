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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Text
{
    using NPOI.Util;
    using NPOI.XDDF.UserModel;
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFBulletStylePicture : IXDDFBulletStyle
    {
        private CT_TextBlipBullet style;
        public XDDFBulletStylePicture(CT_TextBlipBullet style)
        {
            this.style = style;
        }
        public CT_TextBlipBullet GetXmlObject()
        {
            return style;
        }

        public XDDFPicture GetPicture()
        {
            return new XDDFPicture(style.blip);
        }

        public void SetPicture(XDDFPicture picture)
        {
            if(picture != null)
            {
                style.blip = picture.GetXmlObject();
            }
        }
    }
}

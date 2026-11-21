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

    using NPOI.OpenXmlFormats.Dml;

    public enum SchemeColor
    {
        Accent1,
        Accent2,
        Accent3,
        Accent4,
        Accent5,
        Accent6,
        Background1,
        Background2,
        Dark1,
        Dark2,
        FollowedLink,
        Link,
        Light1,
        Light2,
        Placeholder,
        Text1,
        Text2
    }
    public static class SchemeColorExtensions
    {
        private static Dictionary<ST_SchemeColorVal, SchemeColor> reverse = new Dictionary<ST_SchemeColorVal, SchemeColor>()
        {
            { ST_SchemeColorVal.accent1, SchemeColor.Accent1 },
            { ST_SchemeColorVal.accent2, SchemeColor.Accent2 },
            { ST_SchemeColorVal.accent3, SchemeColor.Accent3 },
            { ST_SchemeColorVal.accent4, SchemeColor.Accent4 },
            { ST_SchemeColorVal.accent5, SchemeColor.Accent5 },
            { ST_SchemeColorVal.accent6, SchemeColor.Accent6 },
            { ST_SchemeColorVal.bg1, SchemeColor.Background1 },
            { ST_SchemeColorVal.bg2, SchemeColor.Background2 },
            { ST_SchemeColorVal.dk1, SchemeColor.Dark1 },
            { ST_SchemeColorVal.dk2, SchemeColor.Dark2 },
            { ST_SchemeColorVal.folHlink, SchemeColor.FollowedLink },
            { ST_SchemeColorVal.hlink, SchemeColor.Link },
            { ST_SchemeColorVal.lt1, SchemeColor.Light1 },
            { ST_SchemeColorVal.lt2, SchemeColor.Light2 },
            { ST_SchemeColorVal.phClr, SchemeColor.Placeholder },
            { ST_SchemeColorVal.tx1, SchemeColor.Text1 },
            { ST_SchemeColorVal.tx2, SchemeColor.Text2 },
        };
        public static SchemeColor ValueOf(ST_SchemeColorVal value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_SchemeColorVal", nameof(value));
        }
        public static ST_SchemeColorVal ToST_SchemeColorVal(this SchemeColor value)
        {
            switch(value)
            {
                case SchemeColor.Accent1:
                    return ST_SchemeColorVal.accent1;
                case SchemeColor.Accent2:
                    return ST_SchemeColorVal.accent2;
                case SchemeColor.Accent3:
                    return ST_SchemeColorVal.accent3;
                case SchemeColor.Accent4:
                    return ST_SchemeColorVal.accent4;
                case SchemeColor.Accent5:
                    return ST_SchemeColorVal.accent5;
                case SchemeColor.Accent6:
                    return ST_SchemeColorVal.accent6;
                case SchemeColor.Background1:
                    return ST_SchemeColorVal.bg1;
                case SchemeColor.Background2:
                    return ST_SchemeColorVal.bg2;
                case SchemeColor.Dark1:
                    return ST_SchemeColorVal.dk1;
                case SchemeColor.Dark2:
                    return ST_SchemeColorVal.dk2;
                case SchemeColor.FollowedLink:
                    return ST_SchemeColorVal.folHlink;
                case SchemeColor.Link:
                    return ST_SchemeColorVal.hlink;
                case SchemeColor.Light1:
                    return ST_SchemeColorVal.lt1;
                case SchemeColor.Light2:
                    return ST_SchemeColorVal.lt2;
                case SchemeColor.Placeholder:
                    return ST_SchemeColorVal.phClr;
                case SchemeColor.Text1:
                    return ST_SchemeColorVal.tx1;
                case SchemeColor.Text2:
                    return ST_SchemeColorVal.tx2;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



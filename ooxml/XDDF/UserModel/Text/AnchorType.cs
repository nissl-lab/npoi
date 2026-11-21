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

namespace NPOI.XDDF.UserModel.Text
{

    using NPOI.OpenXmlFormats.Dml;

    public enum AnchorType
    {
        Bottom,
        Center,
        Distributed,
        Justified,
        Top
    }
    public static class AnchorTypeExtensions
    {
        private static Dictionary<ST_TextAnchoringType, AnchorType> reverse = new Dictionary<ST_TextAnchoringType, AnchorType>()
        {
            { ST_TextAnchoringType.b, AnchorType.Bottom },
            { ST_TextAnchoringType.ctr, AnchorType.Center },
            { ST_TextAnchoringType.dist, AnchorType.Distributed },
            { ST_TextAnchoringType.just, AnchorType.Justified },
            { ST_TextAnchoringType.t, AnchorType.Top },
        };
        public static AnchorType ValueOf(ST_TextAnchoringType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextAnchoringType", nameof(value));
        }
        public static ST_TextAnchoringType ToST_TextAnchoringType(this AnchorType value)
        {
            switch(value)
            {
                case AnchorType.Bottom:
                    return ST_TextAnchoringType.b;
                case AnchorType.Center:
                    return ST_TextAnchoringType.ctr;
                case AnchorType.Distributed:
                    return ST_TextAnchoringType.dist;
                case AnchorType.Justified:
                    return ST_TextAnchoringType.just;
                case AnchorType.Top:
                    return ST_TextAnchoringType.t;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



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

    public enum TextAlignment
    {
        Center,
        Distributed,
        Justified,
        JustifiedLow,
        Left,
        Right,
        ThaiDistributed
    }
    public static class TextAlignmentExtensions
    {
        private static Dictionary<ST_TextAlignType, TextAlignment> reverse = new Dictionary<ST_TextAlignType, TextAlignment>()
        {
            { ST_TextAlignType.ctr, TextAlignment.Center },
            { ST_TextAlignType.dist, TextAlignment.Distributed },
            { ST_TextAlignType.just, TextAlignment.Justified },
            { ST_TextAlignType.justLow, TextAlignment.JustifiedLow },
            { ST_TextAlignType.l, TextAlignment.Left },
            { ST_TextAlignType.r, TextAlignment.Right },
            { ST_TextAlignType.thaiDist, TextAlignment.ThaiDistributed },
        };
        public static TextAlignment ValueOf(ST_TextAlignType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextAlignType", nameof(value));
        }
        public static ST_TextAlignType ToST_TextAlignType(this TextAlignment value)
        {
            switch(value)
            {
                case TextAlignment.Center:
                    return ST_TextAlignType.ctr;
                case TextAlignment.Distributed:
                    return ST_TextAlignType.dist;
                case TextAlignment.Justified:
                    return ST_TextAlignType.just;
                case TextAlignment.JustifiedLow:
                    return ST_TextAlignType.justLow;
                case TextAlignment.Left:
                    return ST_TextAlignType.l;
                case TextAlignment.Right:
                    return ST_TextAlignType.r;
                case TextAlignment.ThaiDistributed:
                    return ST_TextAlignType.thaiDist;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



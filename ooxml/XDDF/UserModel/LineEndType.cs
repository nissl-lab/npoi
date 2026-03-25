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

    public enum LineEndType
    {
        Arrow,
        Diamond,
        None,
        Oval,
        Stealth,
        Triangle
    }
    public static class LineEndTypeExtensions
    {
        private static Dictionary<ST_LineEndType, LineEndType> reverse = new Dictionary<ST_LineEndType, LineEndType>()
        {
            { ST_LineEndType.arrow, LineEndType.Arrow },
            { ST_LineEndType.diamond, LineEndType.Diamond },
            { ST_LineEndType.none, LineEndType.None },
            { ST_LineEndType.oval, LineEndType.Oval },
            { ST_LineEndType.stealth, LineEndType.Stealth },
            { ST_LineEndType.triangle, LineEndType.Triangle },
        };
        public static LineEndType ValueOf(ST_LineEndType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_LineEndType", nameof(value));
        }
        public static ST_LineEndType ToST_LineEndType(this LineEndType value)
        {
            switch(value)
            {
                case LineEndType.Arrow:
                    return ST_LineEndType.arrow;
                case LineEndType.Diamond:
                    return ST_LineEndType.diamond;
                case LineEndType.None:
                    return ST_LineEndType.none;
                case LineEndType.Oval:
                    return ST_LineEndType.oval;
                case LineEndType.Stealth:
                    return ST_LineEndType.stealth;
                case LineEndType.Triangle:
                    return ST_LineEndType.triangle;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



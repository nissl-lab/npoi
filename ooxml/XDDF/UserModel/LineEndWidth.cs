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

    public enum LineEndWidth
    {
        Large,
        Medium,
        Small
    }
    public static class LineEndWidthExtensions
    {
        private static Dictionary<ST_LineEndWidth, LineEndWidth> reverse = new Dictionary<ST_LineEndWidth, LineEndWidth>()
        {
            { ST_LineEndWidth.lg, LineEndWidth.Large },
            { ST_LineEndWidth.med, LineEndWidth.Medium },
            { ST_LineEndWidth.sm, LineEndWidth.Small },
        };
        public static LineEndWidth ValueOf(ST_LineEndWidth value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_LineEndWidth", nameof(value));
        }
        public static ST_LineEndWidth ToST_LineEndWidth(this LineEndWidth value)
        {
            switch(value)
            {
                case LineEndWidth.Large:
                    return ST_LineEndWidth.lg;
                case LineEndWidth.Medium:
                    return ST_LineEndWidth.med;
                case LineEndWidth.Small:
                    return ST_LineEndWidth.sm;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



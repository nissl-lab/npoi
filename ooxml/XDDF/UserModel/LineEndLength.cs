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

    public enum LineEndLength
    {
        Large,
        Medium,
        Small
    }
    public static class LineEndLengthExtensions
    {
        private static Dictionary<ST_LineEndLength, LineEndLength> reverse = new Dictionary<ST_LineEndLength, LineEndLength>()
        {
            { ST_LineEndLength.lg, LineEndLength.Large },
            { ST_LineEndLength.med, LineEndLength.Medium },
            { ST_LineEndLength.sm, LineEndLength.Small },
        };
        public static LineEndLength ValueOf(ST_LineEndLength value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_LineEndLength", nameof(value));
        }
        public static ST_LineEndLength ToST_LineEndLength(this LineEndLength value)
        {
            switch(value)
            {
                case LineEndLength.Large:
                    return ST_LineEndLength.lg;
                case LineEndLength.Medium:
                    return ST_LineEndLength.med;
                case LineEndLength.Small:
                    return ST_LineEndLength.sm;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



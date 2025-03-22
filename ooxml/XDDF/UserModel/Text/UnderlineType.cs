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

    public enum UnderlineType
    {
        Dash,
        DashHeavy,
        DashLong,
        DashLongHeavy,
        Double,
        DotDash,
        DotDashHeavy,
        DotDotDash,
        DotDotDashHeavy,
        Dotted,
        DottedHeavy,
        Heavy,
        None,
        Single,
        Wavy,
        WavyDouble,
        WavyHeavy,
        Words
    }
    public static class UnderlineTypeExtensions
    {
        private static Dictionary<ST_TextUnderlineType, UnderlineType> reverse = new Dictionary<ST_TextUnderlineType, UnderlineType>()
        {
            { ST_TextUnderlineType.dash, UnderlineType.Dash },
            { ST_TextUnderlineType.dashHeavy, UnderlineType.DashHeavy },
            { ST_TextUnderlineType.dashLong, UnderlineType.DashLong },
            { ST_TextUnderlineType.dashLongHeavy, UnderlineType.DashLongHeavy },
            { ST_TextUnderlineType.dbl, UnderlineType.Double },
            { ST_TextUnderlineType.dotDash, UnderlineType.DotDash },
            { ST_TextUnderlineType.dotDashHeavy, UnderlineType.DotDashHeavy },
            { ST_TextUnderlineType.dotDotDash, UnderlineType.DotDotDash },
            { ST_TextUnderlineType.dotDotDashHeavy, UnderlineType.DotDotDashHeavy },
            { ST_TextUnderlineType.dotted, UnderlineType.Dotted },
            { ST_TextUnderlineType.dottedHeavy, UnderlineType.DottedHeavy },
            { ST_TextUnderlineType.heavy, UnderlineType.Heavy },
            { ST_TextUnderlineType.none, UnderlineType.None },
            { ST_TextUnderlineType.sng, UnderlineType.Single },
            { ST_TextUnderlineType.wavy, UnderlineType.Wavy },
            { ST_TextUnderlineType.wavyDbl, UnderlineType.WavyDouble },
            { ST_TextUnderlineType.wavyHeavy, UnderlineType.WavyHeavy },
            { ST_TextUnderlineType.words, UnderlineType.Words },
        };
        public static UnderlineType ValueOf(ST_TextUnderlineType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextUnderlineType", nameof(value));
        }
        public static ST_TextUnderlineType ToST_TextUnderlineType(this UnderlineType value)
        {
            switch(value)
            {
                case UnderlineType.Dash:
                    return ST_TextUnderlineType.dash;
                case UnderlineType.DashHeavy:
                    return ST_TextUnderlineType.dashHeavy;
                case UnderlineType.DashLong:
                    return ST_TextUnderlineType.dashLong;
                case UnderlineType.DashLongHeavy:
                    return ST_TextUnderlineType.dashLongHeavy;
                case UnderlineType.Double:
                    return ST_TextUnderlineType.dbl;
                case UnderlineType.DotDash:
                    return ST_TextUnderlineType.dotDash;
                case UnderlineType.DotDashHeavy:
                    return ST_TextUnderlineType.dotDashHeavy;
                case UnderlineType.DotDotDash:
                    return ST_TextUnderlineType.dotDotDash;
                case UnderlineType.DotDotDashHeavy:
                    return ST_TextUnderlineType.dotDotDashHeavy;
                case UnderlineType.Dotted:
                    return ST_TextUnderlineType.dotted;
                case UnderlineType.DottedHeavy:
                    return ST_TextUnderlineType.dottedHeavy;
                case UnderlineType.Heavy:
                    return ST_TextUnderlineType.heavy;
                case UnderlineType.None:
                    return ST_TextUnderlineType.none;
                case UnderlineType.Single:
                    return ST_TextUnderlineType.sng;
                case UnderlineType.Wavy:
                    return ST_TextUnderlineType.wavy;
                case UnderlineType.WavyDouble:
                    return ST_TextUnderlineType.wavyDbl;
                case UnderlineType.WavyHeavy:
                    return ST_TextUnderlineType.wavyHeavy;
                case UnderlineType.Words:
                    return ST_TextUnderlineType.words;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



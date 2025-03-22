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

    public enum PresetLineDash
    {
        Dash,
        DashDot,
        Dot,
        LargeDash,
        LargeDashDot,
        LargeDashDotDot,
        Solid,
        SystemDash,
        SystemDashDot,
        SystemDashDotDot,
        SystemDot
    }
    public static class PresetLineDashExtensions
    {
        private static Dictionary<ST_PresetLineDashVal, PresetLineDash> reverse = new Dictionary<ST_PresetLineDashVal, PresetLineDash>()
        {
            { ST_PresetLineDashVal.dash, PresetLineDash.Dash },
            { ST_PresetLineDashVal.dashDot, PresetLineDash.DashDot },
            { ST_PresetLineDashVal.dot, PresetLineDash.Dot },
            { ST_PresetLineDashVal.lgDash, PresetLineDash.LargeDash },
            { ST_PresetLineDashVal.lgDashDot, PresetLineDash.LargeDashDot },
            { ST_PresetLineDashVal.lgDashDotDot, PresetLineDash.LargeDashDotDot },
            { ST_PresetLineDashVal.solid, PresetLineDash.Solid },
            { ST_PresetLineDashVal.sysDash, PresetLineDash.SystemDash },
            { ST_PresetLineDashVal.sysDashDot, PresetLineDash.SystemDashDot },
            { ST_PresetLineDashVal.sysDashDotDot, PresetLineDash.SystemDashDotDot },
            { ST_PresetLineDashVal.sysDot, PresetLineDash.SystemDot },
        };
        public static PresetLineDash ValueOf(ST_PresetLineDashVal value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_PresetLineDashVal", nameof(value));
        }
        public static ST_PresetLineDashVal ToST_PresetLineDashVal(this PresetLineDash value)
        {
            switch(value)
            {
                case PresetLineDash.Dash:
                    return ST_PresetLineDashVal.dash;
                case PresetLineDash.DashDot:
                    return ST_PresetLineDashVal.dashDot;
                case PresetLineDash.Dot:
                    return ST_PresetLineDashVal.dot;
                case PresetLineDash.LargeDash:
                    return ST_PresetLineDashVal.lgDash;
                case PresetLineDash.LargeDashDot:
                    return ST_PresetLineDashVal.lgDashDot;
                case PresetLineDash.LargeDashDotDot:
                    return ST_PresetLineDashVal.lgDashDotDot;
                case PresetLineDash.Solid:
                    return ST_PresetLineDashVal.solid;
                case PresetLineDash.SystemDash:
                    return ST_PresetLineDashVal.sysDash;
                case PresetLineDash.SystemDashDot:
                    return ST_PresetLineDashVal.sysDashDot;
                case PresetLineDash.SystemDashDotDot:
                    return ST_PresetLineDashVal.sysDashDotDot;
                case PresetLineDash.SystemDot:
                    return ST_PresetLineDashVal.sysDot;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



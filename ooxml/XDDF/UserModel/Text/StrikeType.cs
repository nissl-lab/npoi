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

    public enum StrikeType
    {
        DoubleStrike,
        NoStrike,
        SingleStrike
    }
    public static class StrikeTypeExtensions
    {
        private static Dictionary<ST_TextStrikeType, StrikeType> reverse = new Dictionary<ST_TextStrikeType, StrikeType>()
        {
            { ST_TextStrikeType.dblStrike, StrikeType.DoubleStrike },
            { ST_TextStrikeType.noStrike, StrikeType.NoStrike },
            { ST_TextStrikeType.sngStrike, StrikeType.SingleStrike },
        };
        public static StrikeType ValueOf(ST_TextStrikeType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextStrikeType", nameof(value));
        }
        public static ST_TextStrikeType ToST_TextStrikeType(this StrikeType value)
        {
            switch(value)
            {
                case StrikeType.DoubleStrike:
                    return ST_TextStrikeType.dblStrike;
                case StrikeType.NoStrike:
                    return ST_TextStrikeType.noStrike;
                case StrikeType.SingleStrike:
                    return ST_TextStrikeType.sngStrike;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



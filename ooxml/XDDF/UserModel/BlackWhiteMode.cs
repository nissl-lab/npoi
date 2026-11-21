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

    public enum BlackWhiteMode
    {
        Auto,
        Black,
        BlackGray,
        BlackWhite
    }
    public static class BlackWhiteModeExtensions
    {
        private static Dictionary<ST_BlackWhiteMode, BlackWhiteMode> reverse = new Dictionary<ST_BlackWhiteMode, BlackWhiteMode>()
        {
            { ST_BlackWhiteMode.auto, BlackWhiteMode.Auto },
            { ST_BlackWhiteMode.black, BlackWhiteMode.Black },
            { ST_BlackWhiteMode.blackGray, BlackWhiteMode.BlackGray },
            { ST_BlackWhiteMode.blackWhite, BlackWhiteMode.BlackWhite },
        };
        public static BlackWhiteMode ValueOf(ST_BlackWhiteMode value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_BlackWhiteMode", nameof(value));
        }
        public static ST_BlackWhiteMode ToST_BlackWhiteMode(this BlackWhiteMode value)
        {
            switch(value)
            {
                case BlackWhiteMode.Auto:
                    return ST_BlackWhiteMode.auto;
                case BlackWhiteMode.Black:
                    return ST_BlackWhiteMode.black;
                case BlackWhiteMode.BlackGray:
                    return ST_BlackWhiteMode.blackGray;
                case BlackWhiteMode.BlackWhite:
                    return ST_BlackWhiteMode.blackWhite;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



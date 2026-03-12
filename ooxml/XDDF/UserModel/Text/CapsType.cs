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

    public enum CapsType
    {
        All,
        None,
        Small
    }
    public static class CapsTypeExtensions
    {
        private static Dictionary<ST_TextCapsType, CapsType> reverse = new Dictionary<ST_TextCapsType, CapsType>()
        {
            { ST_TextCapsType.all, CapsType.All },
            { ST_TextCapsType.none, CapsType.None },
            { ST_TextCapsType.small, CapsType.Small },
        };
        public static CapsType ValueOf(ST_TextCapsType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextCapsType", nameof(value));
        }
        public static ST_TextCapsType ToST_TextCapsType(this CapsType value)
        {
            switch(value)
            {
                case CapsType.All:
                    return ST_TextCapsType.all;
                case CapsType.None:
                    return ST_TextCapsType.none;
                case CapsType.Small:
                    return ST_TextCapsType.small;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



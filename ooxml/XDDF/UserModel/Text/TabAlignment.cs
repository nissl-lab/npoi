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

    public enum TabAlignment
    {
        Center,
        Decimal,
        Left,
        Right
    }
    public static class TabAlignmentExtensions
    {
        private static Dictionary<ST_TextTabAlignType, TabAlignment> reverse = new Dictionary<ST_TextTabAlignType, TabAlignment>()
        {
            { ST_TextTabAlignType.ctr, TabAlignment.Center },
            { ST_TextTabAlignType.dec, TabAlignment.Decimal },
            { ST_TextTabAlignType.l, TabAlignment.Left },
            { ST_TextTabAlignType.r, TabAlignment.Right },
        };
        public static TabAlignment ValueOf(ST_TextTabAlignType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextTabAlignType", nameof(value));
        }
        public static ST_TextTabAlignType ToST_TextTabAlignType(this TabAlignment value)
        {
            switch(value)
            {
                case TabAlignment.Center:
                    return ST_TextTabAlignType.ctr;
                case TabAlignment.Decimal:
                    return ST_TextTabAlignType.dec;
                case TabAlignment.Left:
                    return ST_TextTabAlignType.l;
                case TabAlignment.Right:
                    return ST_TextTabAlignType.r;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



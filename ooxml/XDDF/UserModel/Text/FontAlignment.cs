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

    public enum FontAlignment
    {
        Automatic,
        Bottom,
        Baseline,
        Center,
        Top
    }
    public static class FontAlignmentExtensions
    {
        private static Dictionary<ST_TextFontAlignType, FontAlignment> reverse = new Dictionary<ST_TextFontAlignType, FontAlignment>()
        {
            { ST_TextFontAlignType.auto, FontAlignment.Automatic },
            { ST_TextFontAlignType.b, FontAlignment.Bottom },
            { ST_TextFontAlignType.@base, FontAlignment.Baseline },
            { ST_TextFontAlignType.ctr, FontAlignment.Center },
            { ST_TextFontAlignType.t, FontAlignment.Top },
        };
        public static FontAlignment ValueOf(ST_TextFontAlignType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TextFontAlignType", nameof(value));
        }
        public static ST_TextFontAlignType ToST_TextFontAlignType(this FontAlignment value)
        {
            switch(value)
            {
                case FontAlignment.Automatic:
                    return ST_TextFontAlignType.auto;
                case FontAlignment.Bottom:
                    return ST_TextFontAlignType.b;
                case FontAlignment.Baseline:
                    return ST_TextFontAlignType.@base;
                case FontAlignment.Center:
                    return ST_TextFontAlignType.ctr;
                case FontAlignment.Top:
                    return ST_TextFontAlignType.t;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



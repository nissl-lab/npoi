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

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.OpenXmlFormats.Dml.Chart;
    public enum ScatterStyle
    {
        Line,
        LineMarker,
        Marker,
        None,
        Smooth,
        SmoothMarker
    }
    public static class ScatterStyleExtensions
    {
        private static Dictionary<ST_ScatterStyle, ScatterStyle> reverse = new Dictionary<ST_ScatterStyle, ScatterStyle>()
        {
            { ST_ScatterStyle.line, ScatterStyle.Line },
            { ST_ScatterStyle.lineMarker, ScatterStyle.LineMarker },
            { ST_ScatterStyle.marker, ScatterStyle.Marker },
            { ST_ScatterStyle.none, ScatterStyle.None },
            { ST_ScatterStyle.smooth, ScatterStyle.Smooth },
            { ST_ScatterStyle.smoothMarker, ScatterStyle.SmoothMarker },
        };
        public static ScatterStyle ValueOf(ST_ScatterStyle value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_ScatterStyle", nameof(value));
        }
        public static ST_ScatterStyle ToST_ScatterStyle(this ScatterStyle value)
        {
            switch(value)
            {
                case ScatterStyle.Line:
                    return ST_ScatterStyle.line;
                case ScatterStyle.LineMarker:
                    return ST_ScatterStyle.lineMarker;
                case ScatterStyle.Marker:
                    return ST_ScatterStyle.marker;
                case ScatterStyle.None:
                    return ST_ScatterStyle.none;
                case ScatterStyle.Smooth:
                    return ST_ScatterStyle.smooth;
                case ScatterStyle.SmoothMarker:
                    return ST_ScatterStyle.smoothMarker;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}

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
    public enum RadarStyle
    {
        Filled,
        Marker,
        Standard
    }
    public static class RadarStyleExtensions
    {
        private static Dictionary<ST_RadarStyle, RadarStyle> reverse = new Dictionary<ST_RadarStyle, RadarStyle>()
        {
            { ST_RadarStyle.filled, RadarStyle.Filled },
            { ST_RadarStyle.marker, RadarStyle.Marker },
            { ST_RadarStyle.standard, RadarStyle.Standard },
        };
        public static RadarStyle ValueOf(ST_RadarStyle value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_RadarStyle", nameof(value));
        }
        public static ST_RadarStyle ToST_RadarStyle(this RadarStyle value)
        {
            switch(value)
            {
                case RadarStyle.Filled:
                    return ST_RadarStyle.filled;
                case RadarStyle.Marker:
                    return ST_RadarStyle.marker;
                case RadarStyle.Standard:
                    return ST_RadarStyle.standard;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}

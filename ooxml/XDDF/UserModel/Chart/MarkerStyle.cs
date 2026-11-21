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
    public enum MarkerStyle
    {
        Circle,
        Dash,
        Diamond,
        Dot,
        None,
        Picture,
        Plus,
        Square,
        Star,
        Triangle,
        X
    }
    public static class MarkerStyleExtensions
    {
        private static Dictionary<ST_MarkerStyle, MarkerStyle> reverse = new Dictionary<ST_MarkerStyle, MarkerStyle>()
        {
            { ST_MarkerStyle.circle, MarkerStyle.Circle },
            { ST_MarkerStyle.dash, MarkerStyle.Dash },
            { ST_MarkerStyle.diamond, MarkerStyle.Diamond },
            { ST_MarkerStyle.dot, MarkerStyle.Dot },
            { ST_MarkerStyle.none, MarkerStyle.None },
            { ST_MarkerStyle.picture, MarkerStyle.Picture },
            { ST_MarkerStyle.plus, MarkerStyle.Plus },
            { ST_MarkerStyle.square, MarkerStyle.Square },
            { ST_MarkerStyle.star, MarkerStyle.Star },
            { ST_MarkerStyle.triangle, MarkerStyle.Triangle },
            { ST_MarkerStyle.x, MarkerStyle.X },
        };
        public static MarkerStyle ValueOf(ST_MarkerStyle value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_MarkerStyle", nameof(value));
        }
        public static ST_MarkerStyle ToST_MarkerStyle(this MarkerStyle value)
        {
            switch(value)
            {
                case MarkerStyle.Circle:
                    return ST_MarkerStyle.circle;
                case MarkerStyle.Dash:
                    return ST_MarkerStyle.dash;
                case MarkerStyle.Diamond:
                    return ST_MarkerStyle.diamond;
                case MarkerStyle.Dot:
                    return ST_MarkerStyle.dot;
                case MarkerStyle.None:
                    return ST_MarkerStyle.none;
                case MarkerStyle.Picture:
                    return ST_MarkerStyle.picture;
                case MarkerStyle.Plus:
                    return ST_MarkerStyle.plus;
                case MarkerStyle.Square:
                    return ST_MarkerStyle.square;
                case MarkerStyle.Star:
                    return ST_MarkerStyle.star;
                case MarkerStyle.Triangle:
                    return ST_MarkerStyle.triangle;
                case MarkerStyle.X:
                    return ST_MarkerStyle.x;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}
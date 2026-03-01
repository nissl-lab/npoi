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

    public enum RectangleAlignment
    {
        Bottom,
        BottomLeft,
        BottomRight,
        Center,
        Left,
        Right,
        Top,
        TopLeft,
        TopRight
    }
    public static class RectangleAlignmentExtensions
    {
        private static Dictionary<ST_RectAlignment, RectangleAlignment> reverse = new Dictionary<ST_RectAlignment, RectangleAlignment>()
        {
            { ST_RectAlignment.b, RectangleAlignment.Bottom },
            { ST_RectAlignment.bl, RectangleAlignment.BottomLeft },
            { ST_RectAlignment.br, RectangleAlignment.BottomRight },
            { ST_RectAlignment.ctr, RectangleAlignment.Center },
            { ST_RectAlignment.l, RectangleAlignment.Left },
            { ST_RectAlignment.r, RectangleAlignment.Right },
            { ST_RectAlignment.t, RectangleAlignment.Top },
            { ST_RectAlignment.tl, RectangleAlignment.TopLeft },
            { ST_RectAlignment.tr, RectangleAlignment.TopRight },
        };
        public static RectangleAlignment ValueOf(ST_RectAlignment value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_RectAlignment", nameof(value));
        }
        public static ST_RectAlignment ToST_RectAlignment(this RectangleAlignment value)
        {
            switch(value)
            {
                case RectangleAlignment.Bottom:
                    return ST_RectAlignment.b;
                case RectangleAlignment.BottomLeft:
                    return ST_RectAlignment.bl;
                case RectangleAlignment.BottomRight:
                    return ST_RectAlignment.br;
                case RectangleAlignment.Center:
                    return ST_RectAlignment.ctr;
                case RectangleAlignment.Left:
                    return ST_RectAlignment.l;
                case RectangleAlignment.Right:
                    return ST_RectAlignment.r;
                case RectangleAlignment.Top:
                    return ST_RectAlignment.t;
                case RectangleAlignment.TopLeft:
                    return ST_RectAlignment.tl;
                case RectangleAlignment.TopRight:
                    return ST_RectAlignment.tr;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



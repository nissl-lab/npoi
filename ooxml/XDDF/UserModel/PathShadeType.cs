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

    public enum PathShadeType
    {
        Circle,
        Rectangle,
        Shape
    }
    public static class PathShadeTypeExtensions
    {
        private static Dictionary<ST_PathShadeType, PathShadeType> reverse = new Dictionary<ST_PathShadeType, PathShadeType>()
        {
            { ST_PathShadeType.circle, PathShadeType.Circle },
            { ST_PathShadeType.rect, PathShadeType.Rectangle },
            { ST_PathShadeType.shape, PathShadeType.Shape },
        };
        public static PathShadeType ValueOf(ST_PathShadeType value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_PathShadeType", nameof(value));
        }
        public static ST_PathShadeType ToST_PathShadeType(this PathShadeType value)
        {
            switch(value)
            {
                case PathShadeType.Circle:
                    return ST_PathShadeType.circle;
                case PathShadeType.Rectangle:
                    return ST_PathShadeType.rect;
                case PathShadeType.Shape:
                    return ST_PathShadeType.shape;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



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
    public enum AxisPosition
    {
        Bottom,
        Left,
        Right,
        Top
    }
    public static class AxisPositionExtensions
    {
        private static Dictionary<ST_AxPos,AxisPosition> reverse = new Dictionary<ST_AxPos,AxisPosition>()
        {
            { ST_AxPos.b, AxisPosition.Bottom },
            { ST_AxPos.l, AxisPosition.Left },
            { ST_AxPos.r, AxisPosition.Right },
            { ST_AxPos.t, AxisPosition.Top },
        };
        public static AxisPosition ValueOf(ST_AxPos value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_AxPos", nameof(value));
        }
        public static ST_AxPos ToST_AxPos(this AxisPosition value)
        {
            switch(value)
            {
                case AxisPosition.Bottom:
                    return ST_AxPos.b;
                case AxisPosition.Left:
                    return ST_AxPos.l;
                case AxisPosition.Right:
                    return ST_AxPos.r;
                case AxisPosition.Top:
                    return ST_AxPos.t;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


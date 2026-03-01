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
    public enum AxisCrosses
    {
        AutoZero,
        Max,
        Min
    }
    public static class AxisCrossesExtensions
    {
        private static Dictionary<ST_Crosses,AxisCrosses> reverse = new Dictionary<ST_Crosses,AxisCrosses>()
        {
            { ST_Crosses.autoZero, AxisCrosses.AutoZero },
            { ST_Crosses.max, AxisCrosses.Max },
            { ST_Crosses.min, AxisCrosses.Min },
        };
        public static AxisCrosses ValueOf(ST_Crosses value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_Crosses", nameof(value));
        }
        public static ST_Crosses ToST_Crosses(this AxisCrosses value)
        {
            switch(value)
            {
                case AxisCrosses.AutoZero:
                    return ST_Crosses.autoZero;
                case AxisCrosses.Max:
                    return ST_Crosses.max;
                case AxisCrosses.Min:
                    return ST_Crosses.min;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


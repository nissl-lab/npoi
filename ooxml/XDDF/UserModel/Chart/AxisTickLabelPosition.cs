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
    public enum AxisTickLabelPosition
    {
        High,
        Low,
        NextTo,
        None
    }
    public static class AxisTickLabelPositionExtensions
    {
        private static Dictionary<ST_TickLblPos,AxisTickLabelPosition> reverse = new Dictionary<ST_TickLblPos,AxisTickLabelPosition>()
        {
            { ST_TickLblPos.high, AxisTickLabelPosition.High },
            { ST_TickLblPos.low, AxisTickLabelPosition.Low },
            { ST_TickLblPos.nextTo, AxisTickLabelPosition.NextTo },
            { ST_TickLblPos.none, AxisTickLabelPosition.None },
        };
        public static AxisTickLabelPosition ValueOf(ST_TickLblPos value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TickLblPos", nameof(value));
        }
        public static ST_TickLblPos ToST_TickLblPos(this AxisTickLabelPosition value)
        {
            switch(value)
            {
                case AxisTickLabelPosition.High:
                    return ST_TickLblPos.high;
                case AxisTickLabelPosition.Low:
                    return ST_TickLblPos.low;
                case AxisTickLabelPosition.NextTo:
                    return ST_TickLblPos.nextTo;
                case AxisTickLabelPosition.None:
                    return ST_TickLblPos.none;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


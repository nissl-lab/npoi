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
    public enum AxisTickMark
    {
        Cross,
        In,
        None,
        Out
    }
    public static class AxisTickMarkExtensions
    {
        private static Dictionary<ST_TickMark,AxisTickMark> reverse = new Dictionary<ST_TickMark,AxisTickMark>()
        {
            { ST_TickMark.cross, AxisTickMark.Cross },
            { ST_TickMark.@in, AxisTickMark.In },
            { ST_TickMark.none, AxisTickMark.None },
            { ST_TickMark.@out, AxisTickMark.Out },
        };
        public static AxisTickMark ValueOf(ST_TickMark value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TickMark", nameof(value));
        }
        public static ST_TickMark ToST_TickMark(this AxisTickMark value)
        {
            switch(value)
            {
                case AxisTickMark.Cross:
                    return ST_TickMark.cross;
                case AxisTickMark.In:
                    return ST_TickMark.@in;
                case AxisTickMark.None:
                    return ST_TickMark.none;
                case AxisTickMark.Out:
                    return ST_TickMark.@out;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


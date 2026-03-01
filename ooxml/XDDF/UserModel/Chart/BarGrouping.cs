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
    public enum BarGrouping
    {
        Standard,
        Clustered,
        Stacked,
        PercentStacked
    }
    public static class BarGroupingExtensions
    {
        private static Dictionary<ST_BarGrouping, BarGrouping> reverse = new Dictionary<ST_BarGrouping,BarGrouping>()
        {
            { ST_BarGrouping.standard, BarGrouping.Standard },
            { ST_BarGrouping.clustered, BarGrouping.Clustered },
            { ST_BarGrouping.stacked, BarGrouping.Stacked },
            { ST_BarGrouping.percentStacked, BarGrouping.PercentStacked },
        };
        public static BarGrouping ValueOf(ST_BarGrouping value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_BarGrouping", nameof(value));
        }
        public static ST_BarGrouping ToST_BarGrouping(this BarGrouping value)
        {
            switch(value)
            {
                case BarGrouping.Standard:
                    return ST_BarGrouping.standard;
                case BarGrouping.Clustered:
                    return ST_BarGrouping.clustered;
                case BarGrouping.Stacked:
                    return ST_BarGrouping.stacked;
                case BarGrouping.PercentStacked:
                    return ST_BarGrouping.percentStacked;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


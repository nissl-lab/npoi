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
    public enum Grouping
    {
        Standard,
        Stacked,
        PercentStacked
    }
    public static class GroupingExtensions
    {
        private static Dictionary<ST_Grouping, Grouping> reverse = new Dictionary<ST_Grouping,Grouping>()
        {
            { ST_Grouping.standard, Grouping.Standard },
            { ST_Grouping.stacked, Grouping.Stacked },
            { ST_Grouping.percentStacked, Grouping.PercentStacked },
        };
        public static Grouping ValueOf(ST_Grouping value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_Grouping", nameof(value));
        }
        public static ST_Grouping ToST_Grouping(this Grouping value)
        {
            switch(value)
            {
                case Grouping.Standard:
                    return ST_Grouping.standard;
                case Grouping.Stacked:
                    return ST_Grouping.stacked;
                case Grouping.PercentStacked:
                    return ST_Grouping.percentStacked;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}
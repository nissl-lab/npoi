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
    public enum AxisLabelAlignment
    {
        Center,
        Left,
        Right
    }
    public static class AxisLabelAlignmentExtensions
    {
        private static Dictionary<ST_LblAlgn,AxisLabelAlignment> reverse = new Dictionary<ST_LblAlgn,AxisLabelAlignment>()
        {
            { ST_LblAlgn.ctr, AxisLabelAlignment.Center },
            { ST_LblAlgn.l, AxisLabelAlignment.Left },
            { ST_LblAlgn.r, AxisLabelAlignment.Right },
        };
        public static AxisLabelAlignment ValueOf(ST_LblAlgn value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_LblAlgn", nameof(value));
        }
        public static ST_LblAlgn ToST_LblAlgn(this AxisLabelAlignment value)
        {
            switch(value)
            {
                case AxisLabelAlignment.Center:
                    return ST_LblAlgn.ctr;
                case AxisLabelAlignment.Left:
                    return ST_LblAlgn.l;
                case AxisLabelAlignment.Right:
                    return ST_LblAlgn.r;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


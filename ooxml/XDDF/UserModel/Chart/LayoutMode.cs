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

    public enum LayoutMode
    {
        Edge,
        Factor
    }
    public static class LayoutModeExtensions
    {
        private static Dictionary<ST_LayoutMode, LayoutMode> reverse = new Dictionary<ST_LayoutMode, LayoutMode>()
        {
            { ST_LayoutMode.edge, LayoutMode.Edge },
            { ST_LayoutMode.factor, LayoutMode.Factor },
        };
        public static LayoutMode ValueOf(ST_LayoutMode value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_LayoutMode", nameof(value));
        }
        public static ST_LayoutMode ToST_LayoutMode(this LayoutMode value)
        {
            switch(value)
            {
                case LayoutMode.Edge:
                    return ST_LayoutMode.edge;
                case LayoutMode.Factor:
                    return ST_LayoutMode.factor;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}
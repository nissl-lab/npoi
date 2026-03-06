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

    public enum LayoutTarget
    {
        Inner,
        Outer
    }
    public static class LayoutTargetExtensions
    {
        private static Dictionary<ST_LayoutTarget, LayoutTarget> reverse = new Dictionary<ST_LayoutTarget,LayoutTarget>()
        {
            { ST_LayoutTarget.inner, LayoutTarget.Inner },
            { ST_LayoutTarget.outer, LayoutTarget.Outer },
        };
        public static LayoutTarget ValueOf(ST_LayoutTarget value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_LayoutTarget", nameof(value));
        }
        public static ST_LayoutTarget ToST_LayoutTarget(this LayoutTarget value)
        {
            switch(value)
            {
                case LayoutTarget.Inner:
                    return ST_LayoutTarget.inner;
                case LayoutTarget.Outer:
                    return ST_LayoutTarget.outer;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}
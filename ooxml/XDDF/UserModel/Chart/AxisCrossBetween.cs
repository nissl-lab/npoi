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
    public enum AxisCrossBetween
    {
        Between,
        MidpointCategory
    }
    public static class AxisCrossBetweenExtensions
    {
        private static Dictionary<ST_CrossBetween,AxisCrossBetween> reverse = new Dictionary<ST_CrossBetween,AxisCrossBetween>()
        {
            { ST_CrossBetween.between, AxisCrossBetween.Between },
            { ST_CrossBetween.midCat, AxisCrossBetween.MidpointCategory },
        };
        public static AxisCrossBetween ValueOf(ST_CrossBetween value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_CrossBetween", nameof(value));
        }
        public static ST_CrossBetween ToST_CrossBetween(this AxisCrossBetween value)
        {
            switch(value)
            {
                case AxisCrossBetween.Between:
                    return ST_CrossBetween.between;
                case AxisCrossBetween.MidpointCategory:
                    return ST_CrossBetween.midCat;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


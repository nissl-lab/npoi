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
    public enum BarDirection
    {
        Bar,
        Col
    }
    public static class BarDirectionExtensions
    {
        private static Dictionary<ST_BarDir, BarDirection> reverse = new Dictionary<ST_BarDir, BarDirection>()
        {
            { ST_BarDir.bar, BarDirection.Bar },
            { ST_BarDir.col, BarDirection.Col },
        };
        public static BarDirection ValueOf(ST_BarDir value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_BarDir", nameof(value));
        }
        public static ST_BarDir ToST_BarDir(this BarDirection value)
        {
            switch(value)
            {
                case BarDirection.Bar:
                    return ST_BarDir.bar;
                case BarDirection.Col:
                    return ST_BarDir.col;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}


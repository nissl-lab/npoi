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

    public enum DisplayBlanks
    {
        Gap,
        Span,
        Zero
    }
    public static class DisplayBlanksExtensions
    {
        private static Dictionary<ST_DispBlanksAs, DisplayBlanks> reverse = new Dictionary<ST_DispBlanksAs, DisplayBlanks>()
        {
            { ST_DispBlanksAs.gap, DisplayBlanks.Gap },
            { ST_DispBlanksAs.span, DisplayBlanks.Span },
            { ST_DispBlanksAs.zero, DisplayBlanks.Zero },
        };
        public static DisplayBlanks ValueOf(ST_DispBlanksAs value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_DispBlanksAs", nameof(value));
        }
        public static ST_DispBlanksAs ToST_DispBlanksAs(this DisplayBlanks value)
        {
            switch(value)
            {
                case DisplayBlanks.Gap:
                    return ST_DispBlanksAs.gap;
                case DisplayBlanks.Span:
                    return ST_DispBlanksAs.span;
                case DisplayBlanks.Zero:
                    return ST_DispBlanksAs.zero;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}

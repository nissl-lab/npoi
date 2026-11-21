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

namespace NPOI.XDDF.UserModel
{

    using NPOI.OpenXmlFormats.Dml;

    public enum CompoundLine
    {
        Double,
        Single,
        ThickThin,
        ThinThick,
        Triple
    }
    public static class CompoundLineExtensions
    {
        private static Dictionary<ST_CompoundLine, CompoundLine> reverse = new Dictionary<ST_CompoundLine, CompoundLine>()
        {
            { ST_CompoundLine.dbl, CompoundLine.Double },
            { ST_CompoundLine.sng, CompoundLine.Single },
            { ST_CompoundLine.thickThin, CompoundLine.ThickThin },
            { ST_CompoundLine.thinThick, CompoundLine.ThinThick },
            { ST_CompoundLine.tri, CompoundLine.Triple },
        };
        public static CompoundLine ValueOf(ST_CompoundLine value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_CompoundLine", nameof(value));
        }
        public static ST_CompoundLine ToST_CompoundLine(this CompoundLine value)
        {
            switch(value)
            {
                case CompoundLine.Double:
                    return ST_CompoundLine.dbl;
                case CompoundLine.Single:
                    return ST_CompoundLine.sng;
                case CompoundLine.ThickThin:
                    return ST_CompoundLine.thickThin;
                case CompoundLine.ThinThick:
                    return ST_CompoundLine.thinThick;
                case CompoundLine.Triple:
                    return ST_CompoundLine.tri;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



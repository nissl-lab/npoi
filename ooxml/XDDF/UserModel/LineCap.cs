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

    public enum LineCap
    {
        Flat,
        Round,
        Square
    }
    public static class LineCapExtensions
    {
        private static Dictionary<ST_LineCap, LineCap> reverse = new Dictionary<ST_LineCap, LineCap>()
        {
            { ST_LineCap.flat, LineCap.Flat },
            { ST_LineCap.rnd, LineCap.Round },
            { ST_LineCap.sq, LineCap.Square },
        };
        public static LineCap ValueOf(ST_LineCap value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_LineCap", nameof(value));
        }
        public static ST_LineCap ToST_LineCap(this LineCap value)
        {
            switch(value)
            {
                case LineCap.Flat:
                    return ST_LineCap.flat;
                case LineCap.Round:
                    return ST_LineCap.rnd;
                case LineCap.Square:
                    return ST_LineCap.sq;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



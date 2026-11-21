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

    public enum PenAlignment
    {
        Center,
        In
    }
    public static class PenAlignmentExtensions
    {
        private static Dictionary<ST_PenAlignment, PenAlignment> reverse = new Dictionary<ST_PenAlignment, PenAlignment>()
        {
            { ST_PenAlignment.ctr, PenAlignment.Center },
            { ST_PenAlignment.@in, PenAlignment.In },
        };
        public static PenAlignment ValueOf(ST_PenAlignment value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_PenAlignment", nameof(value));
        }
        public static ST_PenAlignment ToST_PenAlignment(this PenAlignment value)
        {
            switch(value)
            {
                case PenAlignment.Center:
                    return ST_PenAlignment.ctr;
                case PenAlignment.In:
                    return ST_PenAlignment.@in;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}



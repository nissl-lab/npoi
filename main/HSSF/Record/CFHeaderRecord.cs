/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Record
{
    using System;

    using NPOI.SS.Util;

    /**
     * Conditional Formatting Header record CFHEADER (0x01B0).
     * Used to describe a {@link CFRuleRecord}.
     * @see CFHeader12Record
     */
    public class CFHeaderRecord : CFHeaderBase, ICloneable
    {
        public static short sid = 0x01B0;

        /** Creates new CFHeaderRecord */
        public CFHeaderRecord()
        {
            CreateEmpty();
        }
        public CFHeaderRecord(CellRangeAddress[] regions, int nRules)
            : base(regions, nRules)
        {
        }

        public CFHeaderRecord(RecordInputStream in1)
        {
            Read(in1);
        }

        protected override string RecordName
        {
            get
            {
                return "CFHEADER";
            }
        }

        public override short Sid
        {
            get{ return sid; }   
        }

        public override object Clone()
        {
            CFHeaderRecord result = new CFHeaderRecord();
            base.CopyTo(result);
            return result;
        }
    }

}

/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Record
{
    using System;
    using System.Collections.Generic;

    /**
     * HorizontalPageBreak record that stores page breaks at rows
     * 
     * This class Is just used so that SID Compares work properly in the RecordFactory
     * @see PageBreakRecord
     * @author Danny Mui (dmui at apache dot org) 
     */
    public class HorizontalPageBreakRecord : PageBreakRecord
    {

        public new const short sid = 0x001B;

        /**
         * 
         */
        public HorizontalPageBreakRecord()
        {
            
        }

        /**
          * @param in the RecordInputstream to Read the record from
          */
        public HorizontalPageBreakRecord(RecordInputStream in1)
            : base(in1)
        {

        }

        /* (non-Javadoc)
         * @see org.apache.poi.hssf.record.Record#Sid
         */
        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            PageBreakRecord result = new HorizontalPageBreakRecord();
            IEnumerator<Break> iterator = GetBreaksEnumerator();
            while (iterator.MoveNext())
            {
                Break original = iterator.Current;
                result.AddBreak(original.main, original.subFrom, original.subTo);
            }
            return result;
        }
    }
}

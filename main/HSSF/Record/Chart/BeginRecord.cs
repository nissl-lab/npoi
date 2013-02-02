
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


namespace NPOI.HSSF.Record.Chart
{

    using System;
    using System.Text;
    using NPOI.Util;



    /**
     * The begin record defines the start of a block of records for a (grpahing
     * data object. This record is matched with a corresponding EndRecord.
     *
     * @see EndRecord
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */

    public class BeginRecord : StandardRecord
    {
        public const short sid = 0x1033;
        public static BeginRecord instance = new BeginRecord();
        public BeginRecord()
        {
        }

        /**
         * Constructs a BeginRecord record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public BeginRecord(RecordInputStream in1)
        {

        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[BEGIN]\n");
            buffer.Append("[/BEGIN]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
        }

        protected override int DataSize
        {
            get { return 0; }
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override Object Clone()
        {
            BeginRecord br = new BeginRecord();
            // No data so nothing to copy
            return br;
        }
    }
}
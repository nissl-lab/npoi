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
    using System.Text;
    using NPOI.Util;

    /**
     * Title: Uncalced Record
     * 
     * If this record occurs in the Worksheet Substream, it indicates that the formulas have not 
     * been recalculated before the document was saved.
     * 
     * @author Olivier Leprince
     */

    public class UncalcedRecord : StandardRecord
    {
        public const short sid = 0x5E;
        private short _reserved;
        /**
         * Default constructor
         */
        public UncalcedRecord()
        {
            _reserved = 0;
        }
        /**
         * Read constructor
         */
        public UncalcedRecord(RecordInputStream in1)
        {
            _reserved = in1.ReadShort();
	    }

        public override short Sid
        {
            get { return sid; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[UNCALCED]\n");
            buffer.Append("[/UNCALCED]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
           out1.WriteShort(_reserved);
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public static int StaticRecordSize
        {
            get { return 6; }
        }
    }
}

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

    using NPOI.Util;
    using System;
    using System.Text;


    /**
     * The end data record is used to denote the end of the subrecords.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class EndSubRecord
       : SubRecord
    {
        public const short sid = 0x00;
        private const int ENCODED_SIZE = 0;

        public EndSubRecord()
        {

        }

        /**
         * Constructs a End record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public EndSubRecord(ILittleEndianInput in1, int size)
        {
            if ((size & 0xFF) != ENCODED_SIZE)
            { // mask out random crap in upper byte
                throw new RecordFormatException("Unexpected size (" + size + ")");
            }

        }

        public override bool IsTerminating
        {
            get
            {
                return true;
            }
        }
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ftEnd]\n");

            buffer.Append("[/ftEnd]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(ENCODED_SIZE);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public override int DataSize
        {
            get { return ENCODED_SIZE; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            EndSubRecord rec = new EndSubRecord();

            return rec;
        }

    }
}
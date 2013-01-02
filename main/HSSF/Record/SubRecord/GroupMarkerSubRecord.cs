
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
     * The Group marker record is used as a position holder for Groups.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class GroupMarkerSubRecord
       : SubRecord
    {
        public const short sid = 0x06;

        private byte[] reserved;    // would really love to know what goes in here.
        private static byte[] EMPTY_BYTE_ARRAY = { };

        public GroupMarkerSubRecord()
        {
            reserved = EMPTY_BYTE_ARRAY;
        }

        /**
         * Constructs a Group marker record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public GroupMarkerSubRecord(ILittleEndianInput in1, int size)
        {
            byte[] buf = new byte[size];
            in1.ReadFully(buf);
            reserved = buf;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            String nl = Environment.NewLine;
            buffer.Append("[ftGmo]" + nl);
            buffer.Append("  reserved = ").Append(HexDump.ToHex(reserved)).Append(nl);
            buffer.Append("[/ftGmo]" + nl);
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(reserved.Length);
            out1.Write(reserved);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public override int DataSize
        {
            get { return reserved.Length; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            GroupMarkerSubRecord rec = new GroupMarkerSubRecord();
            rec.reserved = new byte[reserved.Length];
            for (int i = 0; i < reserved.Length; i++)
                rec.reserved[i] = reserved[i];
            return rec;
        }



    }
}
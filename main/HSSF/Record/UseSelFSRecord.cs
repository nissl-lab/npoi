
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
     * Title:        Use Natural Language Formulas Flag
     * Description:  Tells the GUI if this was written by something that can use
     *               "natural language" formulas. HSSF can't.
     * REFERENCE:  PG 420 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class UseSelFSRecord
       : Record
    {
        public static short sid = 0x160;
        public static short TRUE = 1;
        public static short FALSE = 0;
        private short field_1_flag;

        public UseSelFSRecord()
        {
        }

        /**
         * Constructs a UseSelFS record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public UseSelFSRecord(RecordInputStream in1)
        {
            field_1_flag = in1.ReadShort();
        }

        /**
         * turn the flag on or off
         *
         * @param flag  whether to use natural language formulas or not
         * @see #TRUE
         * @see #FALSE
         */

        public void SetFlag(short flag)
        {
            field_1_flag = flag;
        }

        /**
         * returns whether we use natural language formulas or not
         *
         * @return whether to use natural language formulas or not
         * @see #TRUE
         * @see #FALSE
         */

        public short GetFlag()
        {
            return field_1_flag;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[USESELFS]\n");
            buffer.Append("    .flag            = ")
                .Append(StringUtil.ToHexString(GetFlag())).Append("\n");
            buffer.Append("[/USESELFS]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset,
                                  ((short)0x02));   // 2 bytes (6 total)
            LittleEndian.PutShort(data, 4 + offset, GetFlag());
            return RecordSize;
        }

        public override int RecordSize
        {
            get { return 6; }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}
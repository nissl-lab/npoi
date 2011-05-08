
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
     * Title: Interface Header Record
     * Description: Defines the beginning of Interface records (MMS)
     * REFERENCE:  PG 324 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class InterfaceHdrRecord
       : Record
    {
        public const short sid = 0xe1;
        private short field_1_codepage;   // = 0;

        /**
         * suggested (and probably correct) default
         */

        public static short CODEPAGE = (short)0x4b0;

        public InterfaceHdrRecord()
        {
        }

        /**
         * Constructs an Codepage record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public InterfaceHdrRecord(RecordInputStream in1)
        {
            field_1_codepage = in1.ReadShort();
        }


        public short Codepage
        {
            get { return field_1_codepage; }
            set { field_1_codepage = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[INTERFACEHDR]\n");
            buffer.Append("    .codepage        = ")
                .Append(StringUtil.ToHexString(Codepage)).Append("\n");
            buffer.Append("[/INTERFACEHDR]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset,
                                  ((short)0x02));   // 2 bytes (6 total)
            LittleEndian.PutShort(data, 4 + offset, Codepage);
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

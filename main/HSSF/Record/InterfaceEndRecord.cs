
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
     * Title: Interface End Record
     * Description: Shows where the Interface Records end (MMS)
     *  (has no fields)
     * REFERENCE:  PG 324 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class InterfaceEndRecord
       : Record
    {
        public const short sid = 0xe2;

        public InterfaceEndRecord()
        {
        }

        /**
         * Constructs an InterfaceEnd record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public InterfaceEndRecord(RecordInputStream in1)
        {

        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[INTERFACEEND]\n");
            buffer.Append("[/INTERFACEEND]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset,
                                  ((short)0x00));   // 0 bytes (4 total)
            return RecordSize;
        }

        public override int RecordSize
        {
            get { return 4; }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}

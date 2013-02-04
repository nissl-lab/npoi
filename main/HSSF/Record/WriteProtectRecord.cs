
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
     * Title:        Write Protect Record
     * Description:  Indicated that the sheet/workbook Is Write protected. 
     * REFERENCE:  PG 425 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @version 3.0-pre
     */

    public class WriteProtectRecord
       : StandardRecord
    {
        public const short sid = 0x86;

        public WriteProtectRecord()
        {
        }

        /**
         * Constructs a WriteAccess record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public WriteProtectRecord(RecordInputStream in1)
        {

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[WritePROTECT]\n");
            buffer.Append("[/WritePROTECT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
        }

       protected override int DataSize
       {
           get
           {
               return 0;
           }
       }

        public override short Sid
        {
            get { return sid; }
        }
    }
}
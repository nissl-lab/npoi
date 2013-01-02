
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
     * Title:        Print Headers Record
     * Description:  Whether or not to print the row/column headers when you
     *               enjoy your spReadsheet in the physical form.
     * REFERENCE:  PG 373 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    /* PrintRowCol record */
    public class PrintHeadersRecord
       : StandardRecord
    {
        public const short sid = 0x2a;
        private short field_1_print_headers;

        public PrintHeadersRecord()
        {
        }

        /**
         * Constructs a PrintHeaders record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public PrintHeadersRecord(RecordInputStream in1)
        {
            field_1_print_headers = in1.ReadShort();
        }

        /// <summary>
        /// Get whether to print the headers - y/n
        /// </summary>
        /// <value><c>true</c> if [print headers]; otherwise, <c>false</c>.</value>
        public bool PrintHeaders
        {
            get
            {
                return (field_1_print_headers == 1);
            }
            set 
            {
                field_1_print_headers = (short)((value == true) ?1:0);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PRINTHEADERS]\n");
            buffer.Append("    .printheaders   = ").Append(PrintHeaders)
                .Append("\n");
            buffer.Append("[/PRINTHEADERS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_print_headers);
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            PrintHeadersRecord rec = new PrintHeadersRecord();
            rec.field_1_print_headers = field_1_print_headers;
            return rec;
        }
    }
}
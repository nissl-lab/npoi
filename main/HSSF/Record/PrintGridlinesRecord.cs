
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
     * Title:        Print Gridlines Record
     * Description:  whether to print the gridlines when you enjoy you spReadsheet on paper.
     * REFERENCE:  PG 373 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class PrintGridlinesRecord
       : StandardRecord
    {
        public const short sid = 0x2b;
        private short field_1_print_gridlines;

        public PrintGridlinesRecord()
        {
        }

        /**
         * Constructs a PrintGridlines record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public PrintGridlinesRecord(RecordInputStream in1)
        {
            field_1_print_gridlines = in1.ReadShort();
        }

        /**
         * Get whether or not to print the gridlines (and make your spReadsheet ugly)
         *
         * @return make spReadsheet ugly - Y/N
         */

        public bool PrintGridlines
        {
            get { return (field_1_print_gridlines == 1); }
            set
            {
                if (value)
                {
                    field_1_print_gridlines = 1;
                }
                else
                {
                    field_1_print_gridlines = 0;
                }
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PRINTGRIDLINES]\n");
            buffer.Append("    .printgridlines = ").Append(PrintGridlines)
                .Append("\n");
            buffer.Append("[/PRINTGRIDLINES]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_print_gridlines);
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
            PrintGridlinesRecord rec = new PrintGridlinesRecord();
            rec.field_1_print_gridlines = field_1_print_gridlines;
            return rec;
        }
    }
}
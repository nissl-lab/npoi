
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
     * Title: Function Group Count Record
     * Description:  Number of built in function Groups in the current version of the
     *               SpReadsheet (probably only used on Windoze)
     * REFERENCE:  PG 315 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class FnGroupCountRecord
       : StandardRecord
    {
        public const short sid = 0x9c;

        /**
         * suggested default (14 dec)
         */

        public const short COUNT = 14;
        private short field_1_count;

        public FnGroupCountRecord()
        {
        }

        /**
         * Constructs a FnGroupCount record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public FnGroupCountRecord(RecordInputStream in1)
        {
            field_1_count = in1.ReadShort();
        }
        /**
         * Get the number of built-in functions
         *
         * @return number of built-in functions
         */

        public short Count
        {
            get { return field_1_count; }
            set { field_1_count = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FNGROUPCOUNT]\n");
            buffer.Append("    .count            = ").Append(this.Count)
                .Append("\n");
            buffer.Append("[/FNGROUPCOUNT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Count);
        }

        protected override int DataSize
        {
            get
            {
                return 2;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}
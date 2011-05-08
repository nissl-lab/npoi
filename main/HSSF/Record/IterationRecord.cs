
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
     * Title:        Iteration Record
     * Description:  Tells whether to iterate over forumla calculations or not
     *               (if a formula Is dependant upon another formula's result)
     *               (odd feature for something that can only have 32 elements in
     *                a formula!)
     * REFERENCE:  PG 325 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class IterationRecord
       : Record
    {
        public const short sid = 0x11;
        private short field_1_iteration;

        public IterationRecord()
        {
        }

        /**
         * Constructs an Iteration record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public IterationRecord(RecordInputStream in1)
        {
            field_1_iteration = in1.ReadShort();
        }

        /**
         * Get whether or not to iterate for calculations
         *
         * @return whether iterative calculations are turned off or on
         */

        public bool Iteration
        {
            get { return (field_1_iteration == 1); }
            set
            {
                if (value)
                {
                    field_1_iteration = 1;
                }
                else
                {
                    field_1_iteration = 0;
                }
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[ITERATION]\n");
            buffer.Append("    .iteration      = ").Append(Iteration)
                .Append("\n");
            buffer.Append("[/ITERATION]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset, (short)0x2);
            LittleEndian.PutShort(data, 4 + offset, field_1_iteration);
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

        public override Object Clone()
        {
            IterationRecord rec = new IterationRecord();
            rec.field_1_iteration = field_1_iteration;
            return rec;
        }
    }
}
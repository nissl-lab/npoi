
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
     * Title:        Calc Count Record
     * Description:  Specifies the maximum times the gui should perform a formula
     *               recalculation.  For instance: in the case a formula includes
     *               cells that are themselves a result of a formula and a value
     *               Changes.  This Is essentially a failsafe against an infinate
     *               loop in the event the formulas are not independant. 
     * REFERENCE:  PG 292 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     * @see org.apache.poi.hssf.record.CalcModeRecord
     */

    public class CalcCountRecord
       : StandardRecord
    {
        public const short sid = 0xC;
        private short field_1_iterations;

        public CalcCountRecord()
        {
        }

        /**
         * Constructs a CalcCountRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         *
         */

        public CalcCountRecord(RecordInputStream in1)
        {
            field_1_iterations = in1.ReadShort();
        }

        /**
         * Get the number of iterations to perform
         * @return iterations
         */

        public short Iterations
        {
            get { return field_1_iterations; }
            set { field_1_iterations = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CALCCOUNT]\n");
            buffer.Append("    .iterations     = ")
                .Append(StringUtil.ToHexString(Iterations)).Append("\n");
            buffer.Append("[/CALCCOUNT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Iterations);
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

        public override Object Clone()
        {
            CalcCountRecord rec = new CalcCountRecord();
            rec.field_1_iterations = field_1_iterations;
            return rec;
        }
    }
}
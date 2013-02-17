
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
    using System.Text;
    using NPOI.Util;
    using System;


    /**
     * Title:        Delta Record
     * Description:  controls the accuracy of the calculations
     * REFERENCE:  PG 303 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class DeltaRecord : StandardRecord
    {
        public const short sid = 0x10;
        public const double DEFAULT_VALUE = 0.0010;   // should be .001

        // a double Is an IEEE 8-byte float...damn IEEE and their goofy standards an
        // ambiguous numeric identifiers
        private double field_1_max_change;

        public DeltaRecord(double maxChange)
        {
            field_1_max_change = maxChange;
        }

        /**
         * Constructs a Delta record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public DeltaRecord(RecordInputStream in1)
        {
            field_1_max_change = in1.ReadDouble();
        }


        /**
         * Get the maximum Change
         * @return maxChange - maximum rounding error
         */

        public double MaxChange
        {
            get
            {
                return field_1_max_change;
            }
            set {
                field_1_max_change = value;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DELTA]\n");
            buffer.Append("    .maxChange      = ").Append(MaxChange)
                .Append("\n");
            buffer.Append("[/DELTA]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteDouble(MaxChange);
        }

        protected override int DataSize
        {
            get { return 8; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            return this;
        }
    }
}

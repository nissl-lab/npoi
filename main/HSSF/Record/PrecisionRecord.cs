
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
     * Title:        Precision Record
     * Description:  defines whether to store with full precision or what's Displayed by the gui
     *               (meaning have really screwed up and skewed figures or only think you do!)
     * REFERENCE:  PG 372 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class PrecisionRecord
       : StandardRecord
    {
        public const short sid = 0xE;
        public short field_1_precision;

        public PrecisionRecord()
        {
        }

        /**
         * Constructs a Precision record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public PrecisionRecord(RecordInputStream in1)
        {
            field_1_precision = in1.ReadShort();
        }

        /**
         * Get whether to use full precision or just skew all you figures all to hell.
         *
         * @return fullprecision - or not
         */

        public bool FullPrecision
        {
            get
            {
                return (field_1_precision == 1);
            }
            set
            {
                if (value == true)
                {
                    field_1_precision = 1;
                }
                else
                {
                    field_1_precision = 0;
                }
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PRECISION]\n");
            buffer.Append("    .precision       = ").Append(FullPrecision)
                .Append("\n");
            buffer.Append("[/PRECISION]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_precision);
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

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
    using NPOI.HSSF.Util;


    /**
     * Title:        RK Record
     * Description:  An internal 32 bit number with the two most significant bits
     *               storing the type.  This is part of a bizarre scheme to save disk
     *               space and memory (gee look at all the other whole records that
     *               are in the file just "cause"..,far better to waste Processor
     *               cycles on this then leave on of those "valuable" records out).
     * We support this in Read-ONLY mode.  HSSF Converts these to NUMBER records
     *
     *
     *
     * REFERENCE:  PG 376 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     * @see org.apache.poi.hssf.record.NumberRecord
     */

    public class RKRecord : CellRecord
    {
        public const short sid = 0x27e;
        public const short RK_IEEE_NUMBER = 0;
        public const short RK_IEEE_NUMBER_TIMES_100 = 1;
        public const short RK_INTEGER = 2;
        public const short RK_INTEGER_TIMES_100 = 3;

        private int field_4_rk_number;

        public RKRecord()
        {
        }

        /**
         * Constructs a RK record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public RKRecord(RecordInputStream in1)
            : base(in1)
        {
            field_4_rk_number = in1.ReadInt();
        }

        public int RKField
        {
            get { return field_4_rk_number; }

        }

        /**
         * Get the type of the number
         *
         * @return one of these values:
         *         <OL START="0">
         *             <LI>RK_IEEE_NUMBER</LI>
         *             <LI>RK_IEEE_NUMBER_TIMES_100</LI>
         *             <LI>RK_INTEGER</LI>
         *             <LI>RK_INTEGER_TIMES_100</LI>
         *         </OL>
         */

        public short RKType
        {
            get { return (short)(field_4_rk_number & 3); }
        }

        /**
         * Extract the value of the number
         * 
         * The mechanism for determining the value is dependent on the two
         * low order bits of the raw number. If bit 1 is Set, the number
         * is an integer and can be cast directly as a double, otherwise,
         * it's apparently the exponent and mantissa of a double (and the
         * remaining low-order bits of the double's mantissa are 0's).
         * 
         * If bit 0 is Set, the result of the conversion to a double Is
         * divided by 100; otherwise, the value is left alone.
         * 
         * [Insert picture of Screwy Squirrel in full Napoleonic regalia]
         *
         * @return the value as a proper double (hey, it <B>could</B>
         *         happen)
         */

        public double RKNumber
        {
            get { return RKUtil.DecodeNumber(field_4_rk_number); }
        }


        protected override String RecordName
        {
            get
            {
                return "RK";
            }
        }


        protected override void AppendValueText(StringBuilder sb)
        {
            sb.Append("  .value= ").Append(RKNumber);
        }

        protected override void SerializeValue(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_4_rk_number);
        }


        protected override int ValueDataSize
        {
            get
            {
                return 4;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public new object Clone()
        {
            RKRecord rec = new RKRecord();
            CopyBaseFields(rec);
            rec.field_4_rk_number = field_4_rk_number;
            return rec;
        }
    }
}
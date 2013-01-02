
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

    using NPOI.HSSF.Record;

    /**
     * Specifies the window's zoom magnification.  If this record Isn't present then the windows zoom is 100%. see p384 Excel Dev Kit
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    public class SCLRecord
       : StandardRecord
    {
        public const short sid = 0xa0;
        private short field_1_numerator;
        private short field_2_denominator;


        public SCLRecord()
        {

        }

        /**
         * Constructs a SCL record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SCLRecord(RecordInputStream in1)
        {

            field_1_numerator = in1.ReadShort();
            field_2_denominator = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SCL]\n");
            buffer.Append("    .numerator            = ")
                .Append("0x").Append(HexDump.ToHex(Numerator))
                .Append(" (").Append(Numerator).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .denominator          = ")
                .Append("0x").Append(HexDump.ToHex(Denominator))
                .Append(" (").Append(Denominator).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SCL]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1) {
            out1.WriteShort(field_1_numerator);
            out1.WriteShort(field_2_denominator);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            SCLRecord rec = new SCLRecord();

            rec.field_1_numerator = field_1_numerator;
            rec.field_2_denominator = field_2_denominator;
            return rec;
        }




        /**
         * Get the numerator field for the SCL record.
         */
        public short Numerator
        {
            get
            {
                return field_1_numerator;
            }
            set 
            {
                this.field_1_numerator = value;
            }
        }

        /**
         * Get the denominator field for the SCL record.
         */
        public short Denominator
        {
            get
            {
                return field_2_denominator;
            }
            set 
            {
                this.field_2_denominator = value;
            }
        }


    }
}



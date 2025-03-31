
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
using Cysharp.Text;
    using NPOI.Util;



    /**
     * The Units record describes Units.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class UnitsRecord
       : StandardRecord
    {
        public const short sid = 0x1001;
        private short field_1_units;


        public UnitsRecord()
        {

        }

        /**
         * Constructs a Units record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */
        public UnitsRecord(RecordInputStream in1)
        {
            field_1_units = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[UNITS]\n");
            buffer.Append("    .units                = ")
                .Append("0x").Append(HexDump.ToHex(Units))
                .Append(" (").Append(Units).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/UNITS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_units);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
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
            UnitsRecord rec = new UnitsRecord();

            rec.field_1_units = field_1_units;
            return rec;
        }




        /**
         * Get the Units field for the Units record.
         */
        public short Units
        {
            get
            {
                return field_1_units;
            }
            set 
            {
                this.field_1_units = value;
            }
        }


    }
}


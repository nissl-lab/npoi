
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
        

/*
 * NumberRecord.java
 *
 * Created on October 1, 2001, 8:01 PM
 */
namespace NPOI.HSSF.Record
{

    using NPOI.Util;

    using System;
    using System.Text;
    using NPOI.SS.Util;

    /**
     * Contains a numeric cell value. 
     * REFERENCE:  PG 334 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class NumberRecord :CellRecord
    {
        public const short sid = 0x203;
        private double field_4_value;

        /** Creates new NumberRecord */
        public NumberRecord()
        {
        }

        /**
         * Constructs a Number record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public NumberRecord(RecordInputStream in1):base(in1)
        {
            field_4_value = in1.ReadDouble();
        }

        protected override String RecordName
        {
            get
            {
                return "NUMBER";
            }
        }
        protected override void AppendValueText(StringBuilder sb)
        {
            sb.Append("  .value= ").Append(NumberToTextConverter.ToText(field_4_value));
        }

        protected override void SerializeValue(ILittleEndianOutput out1)
        {
            out1.WriteDouble(Value);
        }
        protected override int ValueDataSize
        {
            get
            {
                return 8;
            }
        }
        /**
         * Get the value for the cell
         *
         * @return double representing the value
         */

        public double Value
        {
            get { return field_4_value; }
            set { field_4_value = value; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override object Clone()
        {
            NumberRecord rec = new NumberRecord();
            CopyBaseFields(rec);
            rec.field_4_value = field_4_value;
            return rec;
        }
    }
}
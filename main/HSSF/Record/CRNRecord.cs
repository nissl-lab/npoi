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

    using NPOI.SS.Formula.Constant;

    /**
     * Title:       CRN  
     * Description: This record stores the contents of an external cell or cell range 
     * REFERENCE:  5.23
     *
     * @author josh micich
     */
    public class CRNRecord : StandardRecord
    {
        public const short sid = 0x5A;

        private int field_1_last_column_index;
        private int field_2_first_column_index;
        private int field_3_row_index;
        private Object[] field_4_constant_values;

        public CRNRecord()
        {
            throw new NotImplementedException("incomplete code");
        }

        public CRNRecord(RecordInputStream in1)
        {
            field_1_last_column_index = in1.ReadByte() & 0x00FF;
            field_2_first_column_index = in1.ReadByte() & 0x00FF;
            field_3_row_index = in1.ReadShort();
            int nValues = field_1_last_column_index - field_2_first_column_index + 1;
            field_4_constant_values = ConstantValueParser.Parse(in1, nValues);
        }


        public int NumberOfCRNs
        {
            get
            {
                return field_1_last_column_index;
            }
        }


        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name).Append(" [CRN");
            sb.Append(" rowIx=").Append(field_3_row_index);
            sb.Append(" firstColIx=").Append(field_2_first_column_index);
            sb.Append(" lastColIx=").Append(field_1_last_column_index);
            sb.Append("]");
            return sb.ToString();
        }
        protected override int DataSize
        {
            get
            {
                return 4 + ConstantValueParser.GetEncodedSize(field_4_constant_values);
            }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(field_1_last_column_index);
            out1.WriteByte(field_2_first_column_index);
            out1.WriteShort(field_3_row_index);
            ConstantValueParser.Encode(out1, field_4_constant_values);
        }

        /**
         * return the non static version of the id for this record.
         */
        public override short Sid
        {
            get { return sid; }
        }
    }

}
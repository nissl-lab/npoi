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
     * XCT ?CRN Count 
     *
     * REFERENCE:  5.114
     *
     * @author Josh Micich
     */
    public class CRNCountRecord : StandardRecord
    {
        public const short sid = 0x59;

        private const short DATA_SIZE = 4;


        private int field_1_number_crn_records;
        private int field_2_sheet_table_index;

        public CRNCountRecord()
        {
            throw new RuntimeException("incomplete code");
        }

        public CRNCountRecord(RecordInputStream in1)
        {
            field_1_number_crn_records = in1.ReadShort();
            if (field_1_number_crn_records < 0)
            {
                // TODO - seems like the sign bit of this field might be used for some other purpose
                // see example file for test case "TestBugs.test19599()"
                field_1_number_crn_records = (short)-field_1_number_crn_records;
            }
            field_2_sheet_table_index = in1.ReadShort();
        }

        public int NumberOfCRNs
        {
            get
            {
                return field_1_number_crn_records;
            }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name).Append(" [XCT");
            sb.Append(" nCRNs=").Append(field_1_number_crn_records);
            sb.Append(" sheetIx=").Append(field_2_sheet_table_index);
            sb.Append("]");
            return sb.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort((short)field_1_number_crn_records);
            out1.WriteShort((short)field_2_sheet_table_index);
        }
        protected override int DataSize
        {
            get
            {
                return DATA_SIZE;
            }
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

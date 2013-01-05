
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
     * Title:        Label SST Record
     * Description:  Refers to a string in the shared string table and Is a column
     *               value.  
     * REFERENCE:  PG 325 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */
    [Serializable]
    public class LabelSSTRecord : CellRecord
    {
        public const short sid = 0xfd;
        private int field_4_sst_index;

        public LabelSSTRecord()
        {
        }

        /**
         * Constructs an LabelSST record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public LabelSSTRecord(RecordInputStream in1)
            : base(in1)
        {
            field_4_sst_index = in1.ReadInt();
        }

        protected override String RecordName
        {
            get
            {
                return "LABELSST";
            }
        }

        /**
         * Get the index to the string in the SSTRecord
         *
         * @return index of string in the SST Table
         * @see org.apache.poi.hssf.record.SSTRecord
         */

        public int SSTIndex
        {
            get { return field_4_sst_index; }
            set { field_4_sst_index = value; }
        }
        protected override void AppendValueText(StringBuilder sb)
        {
            sb.Append("  .sstIndex = ");
            sb.Append(HexDump.ShortToHex(SSTIndex));
        }

        protected override void SerializeValue(ILittleEndianOutput out1)
        {
            out1.WriteInt(SSTIndex);
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

        public override Object Clone()
        {
            LabelSSTRecord rec = new LabelSSTRecord();
            CopyBaseFields(rec);
            rec.field_4_sst_index = field_4_sst_index;
            return rec;
        }
    }
}
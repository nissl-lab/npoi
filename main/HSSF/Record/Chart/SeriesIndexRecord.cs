
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


namespace NPOI.HSSF.Record.Chart
{
    using System;
    using System.Text; 
    using NPOI.Util;


    /**
     * links a series to its position in the series list.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    public class SeriesIndexRecord
       : StandardRecord
    {
        public const short sid = 0x1065;
        private short field_1_index;


        public SeriesIndexRecord()
        {

        }

        /**
         * Constructs a SeriesIndex record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SeriesIndexRecord(RecordInputStream in1)
        {
            field_1_index = in1.ReadShort();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SINDEX]\n");
            buffer.Append("    .index                = ")
                .Append("0x").Append(HexDump.ToHex(Index))
                .Append(" (").Append(Index).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/SINDEX]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_index);
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
            SeriesIndexRecord rec = new SeriesIndexRecord();

            rec.field_1_index = field_1_index;
            return rec;
        }




        /**
         * Get the index field for the SeriesIndex record.
         */
        public short Index
        {
            get { return field_1_index; }
            set { this.field_1_index = value; }
        }


    }
}

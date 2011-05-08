
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
     * Title: Recalc Id Record
     * Description:  This record Contains an ID that marks when a worksheet was last
     *               recalculated. It's an optimization Excel uses to determine if it
     *               needs to  recalculate the spReadsheet when it's opened. So far, only
     *               the two values <c>0xC1 0x01 0x00 0x00 0x80 0x38 0x01 0x00</c>
     *               (do not recalculate) and <c>0xC1 0x01 0x00 0x00 0x60 0x69 0x01
     *               0x00</c> have been seen. If the field <c>isNeeded</c> Is
     *               Set to false (default), then this record Is swallowed during the
     *               serialization Process
     * REFERENCE:  http://chicago.sourceforge.net/devel/docs/excel/biff8.html
     * @author Luc Girardin (luc dot girardin at macrofocus dot com)
     * @version 2.0-pre
     * @see org.apache.poi.hssf.model.Workbook
     */

    public class RecalcIdRecord
       : Record
    {
        public const short sid = 0x1c1;
        public short[] field_1_recalcids;

        private bool isNeeded = true;

        //private bool isNeeded = true;

        public RecalcIdRecord()
        {
        }

        /**
         * Constructs a RECALCID record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public RecalcIdRecord(RecordInputStream in1)
        {
            field_1_recalcids = new short[in1.Remaining / 2];
            for (int k = 0; k < field_1_recalcids.Length; k++)
            {
                field_1_recalcids[k] = in1.ReadShort();
            }
        }

        /**
         * Set the recalc array.
         * @param array of recalc id's
         */

        public void SetRecalcIdArray(short[] array)
        {
            field_1_recalcids = array;
        }

        /**
         * Get the recalc array.
         * @return array of recalc id's
         */

        public short[] GetRecalcIdArray()
        {
            return field_1_recalcids;
        }

        public bool IsNeeded
        {
            get { return  isNeeded; }
            set { this.isNeeded = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[RECALCID]\n");
            buffer.Append("    .elements        = ").Append(field_1_recalcids.Length)
                .Append("\n");
            for (int k = 0; k < field_1_recalcids.Length; k++)
            {
                buffer.Append("    .element_" + k + "       = ")
                    .Append(field_1_recalcids[k]).Append("\n");
            }
            buffer.Append("[/RECALCID]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            short[] tabids = GetRecalcIdArray();
            short Length = (short)(tabids.Length * 2);
            int byteoffset = 4;

            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset,
                                  ((short)Length));

            // 2 (num bytes in a short)
            for (int k = 0; k < (Length / 2); k++)
            {
                LittleEndian.PutShort(data, byteoffset + offset, tabids[k]);
                byteoffset += 2;
            }
            return RecordSize;
        }

        public override int RecordSize
        {
            get { return 4 + (GetRecalcIdArray().Length * 2); }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}

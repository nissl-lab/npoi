
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
     * Title:        Save Recalc Record 
     * Description:  defines whether to recalculate before saving (Set to true)
     * REFERENCE:  PG 381 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class SaveRecalcRecord
       : StandardRecord
    {
        public const short sid = 0x5f;
        private short field_1_recalc;

        public SaveRecalcRecord()
        {
        }

        /**
         * Constructs an SaveRecalc record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public SaveRecalcRecord(RecordInputStream in1)
        {
            field_1_recalc = in1.ReadShort();
        }


        /**
         * Get whether to recalculate formulas/etc before saving or not
         * @return recalc - whether to recalculate or not
         */

        public bool Recalc
        {
            get
            {
                return (field_1_recalc == 1);
            }
            set
            {
                field_1_recalc = (short)((value == true) ? 1
                                            : 0);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SAVERECALC]\n");
            buffer.Append("    .recalc         = ").Append(Recalc)
                .Append("\n");
            buffer.Append("[/SAVERECALC]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_recalc);
        }

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
            SaveRecalcRecord rec = new SaveRecalcRecord();
            rec.field_1_recalc = field_1_recalc;
            return rec;
        }
    }
}
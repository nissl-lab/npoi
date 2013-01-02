
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
     * Title:        RefMode Record
     * Description:  Describes which reference mode to use
     * REFERENCE:  PG 376 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class RefModeRecord
       : StandardRecord
    {
        public const short sid = 0xf;
        public const short USE_A1_MODE = 1;
        public const short USE_R1C1_MODE = 0;
        private short field_1_mode;

        public RefModeRecord()
        {
        }

        /**
         * Constructs a RefMode record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public RefModeRecord(RecordInputStream in1)
        {
            field_1_mode = in1.ReadShort();
        }


        /**
         * Get the reference mode to use (HSSF uses/assumes A1)
         * @return mode to use
         * @see #USE_A1_MODE
         * @see #USE_R1C1_MODE
         */

        public short Mode
        {
            get { return field_1_mode; }
            set { field_1_mode = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[REFMODE]\n");
            buffer.Append("    .mode           = ")
                .Append(StringUtil.ToHexString(Mode)).Append("\n");
            buffer.Append("[/REFMODE]\n");
            return buffer.ToString();
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            RefModeRecord rec = new RefModeRecord();
            rec.field_1_mode = field_1_mode;
            return rec;
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Mode);
        }
    }
}

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
     * Title:        VCenter record
     * Description:  tells whether to center the sheet between vertical margins
     * REFERENCE:  PG 420 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class VCenterRecord
       : StandardRecord
    {
        public const short sid = 0x84;
        private short field_1_vcenter;

        public VCenterRecord()
        {
        }

        /**
         * Constructs a VCENTER record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public VCenterRecord(RecordInputStream in1)
    {
        field_1_vcenter = in1.ReadShort();
    }

        /**
         * Get whether to center vertically or not
         * @return vcenter or not
         */

        public bool VCenter
        {
            get
            {
                return (field_1_vcenter == 1);
            }
            set
            {
                if (value == true)
                {
                    field_1_vcenter = 1;
                }
                else
                {
                    field_1_vcenter = 0;
                }
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[VCENTER]\n");
            buffer.Append("    .vcenter        = ").Append(VCenter)
                .Append("\n");
            buffer.Append("[/VCENTER]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_vcenter);
        }

        protected override int DataSize
        {
            get
            {
                return 2;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            VCenterRecord rec = new VCenterRecord();
            rec.field_1_vcenter = field_1_vcenter;
            return rec;
        }
    }
}
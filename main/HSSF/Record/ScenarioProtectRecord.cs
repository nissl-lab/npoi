
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
     * Title: Scenario Protect Record
     * Description:  I have no idea what a Scenario is or why on would want to 
     * protect it with the lamest "security" ever invented.  However this record tells
     * excel "I want to protect my scenarios" (0xAF) with lame security.  It appears 
     * in conjunction with the PASSWORD and PROTECT records as well as its object 
     * protect cousin.
     * REFERENCE:  PG 383 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     */

    public class ScenarioProtectRecord
       : StandardRecord
    {
        public const short sid = 0xdd;
        private short field_1_protect;

        public ScenarioProtectRecord()
        {
        }

        /**
         * Constructs a Protect record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public ScenarioProtectRecord(RecordInputStream in1)
        {
            field_1_protect = in1.ReadShort();
        }

        /**
         * Get whether the sheet is protected or not
         * @return whether to protect the sheet or not
         */

        public bool Protect
        {
            get { return (field_1_protect == 1); }
            set
            {
                if (value)
                {
                    field_1_protect = 1;
                }
                else
                {
                    field_1_protect = 0;
                }
            }
        }
        public override short Sid
        {
            get
            {
                return sid;
            }
        }
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SCENARIOPROTECT]\n");
            buffer.Append("    .protect         = ").Append(Protect)
                .Append("\n");
            buffer.Append("[/SCENARIOPROTECT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_protect);
        }

        protected override int DataSize
        {
            get
            {
                return 2;
            }
        }

        public override Object Clone()
        {
            ScenarioProtectRecord rec = new ScenarioProtectRecord();
            rec.field_1_protect = field_1_protect;
            return rec;
        }
    }
}
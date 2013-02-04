
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
     * Title:        Country Record (aka WIN.INI country)
     * Description:  used for localization.  Currently HSSF always Sets this to 1
     * and it seems to work fine even in Germany.  (es geht's auch fuer Deutschland)
     *
     * REFERENCE:  PG 298 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class CountryRecord : StandardRecord
    {
        public const short sid = 0x8c;

        // 1 for US
        private short field_1_default_country;
        private short field_2_current_country;

        public CountryRecord()
        {
        }

        /**
         * Constructs a CountryRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public CountryRecord(RecordInputStream in1)
        {
            field_1_default_country = in1.ReadShort();
            field_2_current_country = in1.ReadShort();
        }

        /**
         * Gets the default country
         *
         * @return country ID (1 = US)
         */

        public short DefaultCountry
        {
            get
            {
                return field_1_default_country;
            }
            set { field_1_default_country = value; }
        }

        /**
         * Gets the current country
         *
         * @return country ID (1 = US)
         */

        public short CurrentCountry
        {
            get
            {
                return field_2_current_country;
            }
            set { field_2_current_country = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[COUNTRY]\n");
            buffer.Append("    .defaultcountry  = ")
                .Append(StringUtil.ToHexString(DefaultCountry)).Append("\n");
            buffer.Append("    .currentcountry  = ")
                .Append(StringUtil.ToHexString(CurrentCountry)).Append("\n");
            buffer.Append("[/COUNTRY]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(DefaultCountry);
            out1.WriteShort(CurrentCountry);
        }

        protected override int DataSize
        {
            get { return 4; }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}
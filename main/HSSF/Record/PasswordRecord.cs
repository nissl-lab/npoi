
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
     * Title:        Password Record
     * Description:  stores the encrypted password for a sheet or workbook (HSSF doesn't support encryption)
     * REFERENCE:  PG 371 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class PasswordRecord : StandardRecord
    {
        public const short sid = 0x13;
        private int field_1_password;   // not sure why this Is only 2 bytes, but it Is... go figure

        public PasswordRecord(int password)
        {
            field_1_password = password;
        }

        /**
         * Constructs a Password record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public PasswordRecord(RecordInputStream in1)
        {
            field_1_password = in1.ReadShort();
        }

        //this Is the world's lamest "security".  thanks to Wouter van Vugt for making me
        //not have to try real hard.  -ACO
        public static short HashPassword(String password)
        {
            byte[] passwordChars = Encoding.UTF8.GetBytes(password);
            int hash = 0;
            if (passwordChars.Length > 0)
            {
                int charIndex = passwordChars.Length;
                while (charIndex-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= passwordChars[charIndex];
                }
                // also hash with charcount
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordChars.Length;
                hash ^= (0x8000 | ('N' << 8) | 'K');
            }
            return (short)hash;
        }

        /**
         * Get the password
         *
         * @return short  representing the password
         */
        public int Password
        {
            get { return field_1_password; }
            set { field_1_password = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PASSWORD]\n");
            buffer.Append("    .password       = ")
                .Append(StringUtil.ToHexString(Password)).Append("\n");
            buffer.Append("[/PASSWORD]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_password);
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        /**
         * Clone this record.
         */
        public override Object Clone()
        {
            return new PasswordRecord(field_1_password);
        }

    }
}
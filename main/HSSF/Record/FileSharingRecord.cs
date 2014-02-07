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
     * Title:        FILESHARING
     * Description:  stores the encrypted Readonly for a workbook (Write protect) 
     * This functionality Is accessed from the options dialog box available when performing 'Save As'.<p/>
     * REFERENCE:  PG 314 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)<p/>
     * @author Andrew C. Oliver (acoliver at apache dot org)
     */
    public class FileSharingRecord : StandardRecord
    {

        public const short sid = 0x5b;
        private short field_1_Readonly;
        private short field_2_password;
        private byte field_3_username_unicode_options;
        private String field_3_username_value;

        public FileSharingRecord() { }


        /**
         * Constructs a FileSharing record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public FileSharingRecord(RecordInputStream in1)
        {
            field_1_Readonly = in1.ReadShort();
            field_2_password = in1.ReadShort();

            int nameLen = in1.ReadShort();

            if (nameLen > 0)
            {
                // TODO - Current examples(3) from junits only have zero Length username. 
                field_3_username_unicode_options = (byte)in1.ReadByte();
                field_3_username_value = in1.ReadCompressedUnicode(nameLen);
                
                if (field_3_username_value == null)
                {
                   // In some cases the user name can be null after reading from
                   // the input stream so we make sure this has a value
                   field_3_username_value = "";
                }
            }
            else
            {
                field_3_username_value = "";
            }
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
         * Get the Readonly
         *
         * @return short  representing if this Is Read only (1 = true)
         */
        public short ReadOnly
        {
            get
            {
                return field_1_Readonly;
            }
            set { field_1_Readonly = value; }
        }

        /**
         * @returns password hashed with hashPassword() (very lame)
         */
        public short Password
        {
            get
            {
                return field_2_password;
            }
            set { field_2_password = value; }
        }


        /**
         * @returns username of the user that Created the file
         */
        public String Username
        {
            get { return field_3_username_value; }
            set { field_3_username_value = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FILESHARING]\n");
            buffer.Append("    .Readonly       = ")
                .Append(ReadOnly == 1 ? "true" : "false").Append("\n");
            buffer.Append("    .password       = ")
                .Append(StringUtil.ToHexString(Password)).Append("\n");
            buffer.Append("    .username       = ")
                .Append(Username).Append("\n");
            buffer.Append("[/FILESHARING]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            // TODO - junit
            out1.WriteShort(ReadOnly);
            out1.WriteShort(Password);
            out1.WriteShort(field_3_username_value.Length);
            if (field_3_username_value.Length > 0)
            {
                out1.WriteByte(field_3_username_unicode_options);
                StringUtil.PutCompressedUnicode(Username, out1);
            }
        }

        protected override int DataSize
        {
            get
            {
                int nameLen = field_3_username_value.Length;
                if (nameLen < 1)
                {
                    return 6;
                }
                return 7 + nameLen;
            }
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        /**
         * Clone this record.
         */
        public override Object Clone()
        {
            FileSharingRecord Clone = new FileSharingRecord();
            Clone.ReadOnly = field_1_Readonly;
            Clone.Password = field_2_password;
            Clone.Username = field_3_username_value;
            return Clone;
        }
    }
}

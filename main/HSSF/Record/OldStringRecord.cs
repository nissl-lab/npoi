/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * Biff2 - Biff 4 Label Record (0x0007 / 0x0207) - read only support for 
     *  formula string results.
     */
    public class OldStringRecord
    {
        public const short biff2_sid = 0x0007;
        public const short biff345_sid = 0x0207;

        private short sid;
        private short field_1_string_len;
        private byte[] field_2_bytes;
        private CodepageRecord codepage;

        /**
         * @param in the RecordInputstream to read the record from
         */
        public OldStringRecord(RecordInputStream in1)
        {
            sid = in1.Sid;

            if (in1.Sid == biff2_sid)
            {
                field_1_string_len = (short)in1.ReadUByte();
            }
            else
            {
                field_1_string_len = in1.ReadShort();
            }

            // Can only decode properly later when you know the codepage
            field_2_bytes = new byte[field_1_string_len];
            in1.Read(field_2_bytes, 0, field_1_string_len);
        }

        public bool IsBiff2
        {
            get
            {
                return sid == biff2_sid;
            }
        }

        public short Sid
        {
            get
            {
                return sid;
            }
        }

        public void SetCodePage(CodepageRecord codepage)
        {
            this.codepage = codepage;
        }

        /**
         * @return The string represented by this record.
         */
        public String GetString()
        {
            return GetString(field_2_bytes, codepage);
        }

        protected internal static String GetString(byte[] data, CodepageRecord codepage)
        {
            int cp = CodePageUtil.CP_ISO_8859_1;
            if (codepage != null)
            {
                cp = codepage.Codepage & 0xffff;
            }
            try
            {
                return CodePageUtil.GetStringFromCodePage(data, cp);
            }
            catch (EncoderFallbackException uee)
            {
                throw new ArgumentException("Unsupported codepage requested", uee);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[OLD STRING]\n");
            buffer.Append("    .string            = ")
                .Append(GetString()).Append("\n");
            buffer.Append("[/OLD STRING]\n");
            return buffer.ToString();
        }
    }

}
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

namespace NPOI.HSSF.Record.Common
{
    using System;

    using NPOI.HSSF.Record;
    using NPOI.Util;
    using System.Text;

    /**
     * Title: FeatProtection (Protection Shared Feature) common record part
     * 
     * This record part specifies Protection data for a sheet, stored
     *  as part of a Shared Feature. It can be found in records such
     *  as {@link FeatRecord}
     */
    public class FeatProtection : SharedFeature
    {
        public const long NO_SELF_RELATIVE_SECURITY_FEATURE = 0;
        public const long HAS_SELF_RELATIVE_SECURITY_FEATURE = 1;

        private int fSD;

        /**
         * 0 means no password. Otherwise indicates the
         *  password verifier algorithm (same kind as 
         *   {@link PasswordRecord} and
         *   {@link PasswordRev4Record})
         */
        private int passwordVerifier;

        private String title;
        private byte[] securityDescriptor;

        public FeatProtection()
        {
            securityDescriptor = new byte[0];
        }

        public FeatProtection(RecordInputStream in1)
        {
            fSD = in1.ReadInt();
            passwordVerifier = in1.ReadInt();

            title = StringUtil.ReadUnicodeString(in1);

            securityDescriptor = in1.ReadRemainder();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append(" [FEATURE PROTECTION]\n");
            buffer.Append("   Self Relative = " + fSD);
            buffer.Append("   Password Verifier = " + passwordVerifier);
            buffer.Append("   Title = " + title);
            buffer.Append("   Security Descriptor Size = " + securityDescriptor.Length);
            buffer.Append(" [/FEATURE PROTECTION]\n");
            return buffer.ToString();
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(fSD);
            out1.WriteInt(passwordVerifier);
            StringUtil.WriteUnicodeString(out1, title);
            out1.Write(securityDescriptor);
        }

        public int DataSize
        {
            get
            {
                return 4 + 4 + StringUtil.GetEncodedSize(title) + securityDescriptor.Length;
            }
        }

        public int GetPasswordVerifier()
        {
            return passwordVerifier;
        }
        public void SetPasswordVerifier(int passwordVerifier)
        {
            this.passwordVerifier = passwordVerifier;
        }

        public String GetTitle()
        {
            return title;
        }
        public void SetTitle(String title)
        {
            this.title = title;
        }

        public int GetFSD()
        {
            return fSD;
        }
    }
}
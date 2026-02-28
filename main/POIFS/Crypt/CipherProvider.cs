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

namespace NPOI.POIFS.Crypt
{
    using System;

    public class CipherProvider
    {
        public static readonly CipherProvider rc4 = new CipherProvider("RC4", 1,
            "Microsoft Base Cryptographic Provider v1.0");
        public static readonly CipherProvider aes = new CipherProvider("AES", 0x18,
            "Microsoft Enhanced RSA and AES Cryptographic Provider");

        public static CipherProvider[] Values = new CipherProvider[] { rc4, aes };

        public static CipherProvider FromEcmaId(int ecmaId)
        {
            foreach (CipherProvider cp in CipherProvider.Values)
            {
                if (cp.ecmaId == ecmaId) return cp;
            }
            throw new EncryptedDocumentException("cipher provider not found");
        }

        public String jceId { get;set;}
        public int ecmaId { get;set;}
        public String cipherProviderName { get;set;}

        public CipherProvider(String jceId, int ecmaId, String cipherProviderName)
        {
            this.jceId = jceId;
            this.ecmaId = ecmaId;
            this.cipherProviderName = cipherProviderName;
        }
    }
}

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

namespace NPOI.POIFS.Crypt.CryptoAPI
{
    using System;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.Standard;
    using NPOI.Util;

    public class CryptoAPIEncryptionHeader : StandardEncryptionHeader
    {

        public CryptoAPIEncryptionHeader(ILittleEndianInput is1) : base(is1)
        {
        }

        protected internal CryptoAPIEncryptionHeader(CipherAlgorithm cipherAlgorithm,
                HashAlgorithm hashAlgorithm, int keyBits, int blockSize,
                ChainingMode chainingMode)
            : base(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode)
        {
            
        }

        public void SetKeySize(int keyBits)
        {
            // Microsoft Base Cryptographic Provider is limited up to 40 bits
            // http://msdn.microsoft.com/en-us/library/windows/desktop/aa375599(v=vs.85).aspx
            bool found = false;
            foreach (int size in CipherAlgorithm.allowedKeySize)
            {
                if (size == keyBits)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                throw new EncryptedDocumentException("invalid keysize " + keyBits + " for cipher algorithm " + CipherAlgorithm);
            }
            base.KeySize = (keyBits);
            if (keyBits > 40)
            {
                CspName = ("Microsoft Enhanced Cryptographic Provider v1.0");
            }
            else
            {
                CspName = (CipherProvider.rc4.cipherProviderName);
            }
        }
    }

}
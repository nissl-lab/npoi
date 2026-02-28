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

    public class CryptoAPIEncryptionVerifier : StandardEncryptionVerifier {

        protected internal CryptoAPIEncryptionVerifier(ILittleEndianInput is1,
                CryptoAPIEncryptionHeader header) :
            base(is1, header)
        {
        
        }

        protected internal CryptoAPIEncryptionVerifier(CipherAlgorithm cipherAlgorithm,
                HashAlgorithm hashAlgorithm, int keyBits, int blockSize,
                ChainingMode chainingMode)
            : base(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode)
        { 
            
        }

        protected new void SetSalt(byte[] salt) {
            base.SetSalt(salt);
        }

        protected new void SetEncryptedVerifier(byte[] encryptedVerifier) {
            base.SetEncryptedVerifier(encryptedVerifier);
        }

        protected new void SetEncryptedVerifierHash(byte[] encryptedVerifierHash) {
            base.SetEncryptedVerifierHash(encryptedVerifierHash);
        }
    }
}

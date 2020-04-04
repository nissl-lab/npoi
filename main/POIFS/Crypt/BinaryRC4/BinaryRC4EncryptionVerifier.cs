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

namespace NPOI.POIFS.Crypt.BinaryRC4
{
    using System;
    using System.Diagnostics;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.Standard;
    using NPOI.Util;

    public class BinaryRC4EncryptionVerifier : EncryptionVerifier, EncryptionRecord
    {

        protected internal BinaryRC4EncryptionVerifier()
        {
            SpinCount = (-1);
            CipherAlgorithm = (CipherAlgorithm.rc4);
            ChainingMode = (null);
            EncryptedKey = (null);
            HashAlgorithm = (HashAlgorithm.md5);
        }

        protected internal BinaryRC4EncryptionVerifier(ILittleEndianInput is1)
        {
            byte[] salt = new byte[16];
            is1.ReadFully(salt);
            SetSalt(salt);
            byte[] encryptedVerifier = new byte[16];
            is1.ReadFully(encryptedVerifier);
            EncryptedVerifier = (encryptedVerifier);
            byte[] encryptedVerifierHash = new byte[16];
            is1.ReadFully(encryptedVerifierHash);
            EncryptedVerifierHash = (encryptedVerifierHash);
            SpinCount = (-1);
            CipherAlgorithm = (CipherAlgorithm.rc4);
            ChainingMode = (null);
            EncryptedKey = (null);
            HashAlgorithm = (HashAlgorithm.md5);
        }

        protected internal void SetSalt(byte[] salt)
        {
            if (salt == null || salt.Length != 16)
            {
                throw new EncryptedDocumentException("invalid verifier salt");
            }

            base.Salt = (salt);
        }

        //protected internal void SetEncryptedVerifier(byte[] encryptedVerifier)
        //{
        //    base.EncryptedVerifier = (encryptedVerifier);
        //}

        //protected internal void SetEncryptedVerifierHash(byte[] encryptedVerifierHash)
        //{
        //    base.EncryptedVerifierHash = (encryptedVerifierHash);
        //}

        public void Write(LittleEndianByteArrayOutputStream bos)
        {
            byte[] salt = Salt;
            Debug.Assert(salt.Length == 16);
            bos.Write(salt);
            byte[] encryptedVerifier = EncryptedVerifier;
            Debug.Assert(encryptedVerifier.Length == 16);
            bos.Write(encryptedVerifier);
            byte[] encryptedVerifierHash = EncryptedVerifierHash;
            Debug.Assert(encryptedVerifierHash.Length == 16);
            bos.Write(encryptedVerifierHash);
        }

    }
}

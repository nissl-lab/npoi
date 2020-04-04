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
namespace NPOI.POIFS.Crypt.Standard
{
    using System;
    using System.Diagnostics;
    using NPOI.POIFS.Crypt;
    using NPOI.Util;

    /**
     * Used when Checking if a key is valid for a document 
     */
    public class StandardEncryptionVerifier : EncryptionVerifier , EncryptionRecord {
        private static int SPIN_COUNT = 50000;
        private int verifierHashSize;

        internal StandardEncryptionVerifier(ILittleEndianInput is1, StandardEncryptionHeader header) {
            int saltSize = is1.ReadInt();

            if (saltSize != 16) {
                throw new Exception("Salt size != 16 !?");
            }

            byte[] salt = new byte[16];
            is1.ReadFully(salt);
            SetSalt(salt);

            byte[] encryptedVerifier = new byte[16];
            is1.ReadFully(encryptedVerifier);
            SetEncryptedVerifier(encryptedVerifier);

            verifierHashSize = is1.ReadInt();

            byte[] encryptedVerifierHash = new byte[header.CipherAlgorithm.encryptedVerifierHashLength];
            is1.ReadFully(encryptedVerifierHash);
            SetEncryptedVerifierHash(encryptedVerifierHash);

            SpinCount = (SPIN_COUNT);
            CipherAlgorithm = (header.CipherAlgorithm);
            ChainingMode = (header.ChainingMode);
            EncryptedKey = (null);
            HashAlgorithm = (header.HashAlgorithm);
        }

        protected internal StandardEncryptionVerifier(CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm, int keyBits, int blockSize, ChainingMode chainingMode) {
            CipherAlgorithm = (cipherAlgorithm);
            HashAlgorithm = (hashAlgorithm);
            ChainingMode = (chainingMode);
            SpinCount = (SPIN_COUNT);
            verifierHashSize = hashAlgorithm.hashSize;
        }

        // make method visible for this package
        protected internal void SetSalt(byte[] salt) {
            if (salt == null || salt.Length != 16) {
                throw new EncryptedDocumentException("invalid verifier salt");
            }
            base.Salt = (salt);
        }

        // make method visible for this package
        protected internal void SetEncryptedVerifier(byte[] encryptedVerifier) {
            base.EncryptedVerifier = (encryptedVerifier);
        }

        // make method visible for this package
        protected internal void SetEncryptedVerifierHash(byte[] encryptedVerifierHash) {
            base.EncryptedVerifierHash = (encryptedVerifierHash);
        }

        public void Write(LittleEndianByteArrayOutputStream bos) {
            // see [MS-OFFCRYPTO] - 2.3.4.9
            byte[] salt = Salt;
            Debug.Assert(salt.Length == 16);
            bos.WriteInt(salt.Length); // salt size
            bos.Write(salt);

            // The resulting Verifier value MUST be an array of 16 bytes.
            byte[] encryptedVerifier = EncryptedVerifier;
            Debug.Assert(encryptedVerifier.Length == 16);
            bos.Write(encryptedVerifier);

            // The number of bytes used by the decrypted Verifier hash is given by
            // the VerifierHashSize field, which MUST be 20
            bos.WriteInt(20);

            // EncryptedVerifierHash: An array of bytes that Contains the encrypted form of the hash of
            // the randomly generated Verifier value. The length of the array MUST be the size of the
            // encryption block size multiplied by the number of blocks needed to encrypt the hash of the
            // Verifier. If the encryption algorithm is RC4, the length MUST be 20 bytes. If the encryption
            // algorithm is AES, the length MUST be 32 bytes. After decrypting the EncryptedVerifierHash
            // field, only the first VerifierHashSize bytes MUST be used.
            byte[] encryptedVerifierHash = EncryptedVerifierHash;
            Debug.Assert(encryptedVerifierHash.Length == CipherAlgorithm.encryptedVerifierHashLength);
            bos.Write(encryptedVerifierHash);
        }

        protected int GetVerifierHashSize() {
            return verifierHashSize;
        }
    }
}
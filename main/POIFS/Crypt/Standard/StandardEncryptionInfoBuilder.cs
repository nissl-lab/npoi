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
    using NPOI.POIFS.Crypt;
    using NPOI.Util;

    public class StandardEncryptionInfoBuilder : IEncryptionInfoBuilder
    {

        EncryptionInfo info;
        StandardEncryptionHeader header;
        StandardEncryptionVerifier verifier;
        StandardDecryptor decryptor;
        StandardEncryptor encryptor;

        /**
         * Initialize the builder from a stream
         */
        public void Initialize(EncryptionInfo info, ILittleEndianInput dis)
        {
            this.info = info;

            int hSize = dis.ReadInt();
            header = new StandardEncryptionHeader(dis);
            verifier = new StandardEncryptionVerifier(dis, header);

            if (info.VersionMinor == 2 && (info.VersionMajor == 3 || info.VersionMajor == 4))
            {
                decryptor = new StandardDecryptor(this);
            }
        }

        /**
         * Initialize the builder from scratch
         */
        public void Initialize(EncryptionInfo info, CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm, int keyBits, int blockSize, ChainingMode chainingMode)
        {
            this.info = info;

            if (cipherAlgorithm == null)
            {
                cipherAlgorithm = CipherAlgorithm.aes128;
            }
            if (cipherAlgorithm != CipherAlgorithm.aes128 &&
                cipherAlgorithm != CipherAlgorithm.aes192 &&
                cipherAlgorithm != CipherAlgorithm.aes256)
            {
                throw new EncryptedDocumentException("Standard encryption only supports AES128/192/256.");
            }

            if (hashAlgorithm == null)
            {
                hashAlgorithm = HashAlgorithm.sha1;
            }
            if (hashAlgorithm != HashAlgorithm.sha1)
            {
                throw new EncryptedDocumentException("Standard encryption only supports SHA-1.");
            }
            if (chainingMode == null)
            {
                chainingMode = ChainingMode.ecb;
            }
            if (chainingMode != ChainingMode.ecb)
            {
                throw new EncryptedDocumentException("Standard encryption only supports ECB chaining.");
            }
            if (keyBits == -1)
            {
                keyBits = cipherAlgorithm.defaultKeySize;
            }
            if (blockSize == -1)
            {
                blockSize = cipherAlgorithm.blockSize;
            }
            bool found = false;
            foreach (int ks in cipherAlgorithm.allowedKeySize)
            {
                found |= (ks == keyBits);
            }
            if (!found)
            {
                throw new EncryptedDocumentException("KeySize " + keyBits + " not allowed for Cipher " + cipherAlgorithm.ToString());
            }
            header = new StandardEncryptionHeader(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode);
            verifier = new StandardEncryptionVerifier(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode);
            decryptor = new StandardDecryptor(this);
            encryptor = new StandardEncryptor(this);
        }

        public StandardEncryptionHeader GetHeader()
        {
            return header;
        }

        public StandardEncryptionVerifier GetVerifier()
        {
            return verifier;
        }

        public StandardDecryptor GetDecryptor()
        {
            return decryptor;
        }

        public StandardEncryptor GetEncryptor()
        {
            return encryptor;
        }

        public EncryptionInfo GetEncryptionInfo()
        {
            return info;
        }

        EncryptionHeader IEncryptionInfoBuilder.GetHeader()
        {
            return GetHeader();
        }

        EncryptionVerifier IEncryptionInfoBuilder.GetVerifier()
        {
            return GetVerifier();
        }

        Decryptor IEncryptionInfoBuilder.GetDecryptor()
        {
            return GetDecryptor();
        }

        Encryptor IEncryptionInfoBuilder.GetEncryptor()
        {
            return GetEncryptor();
        }
    }

}
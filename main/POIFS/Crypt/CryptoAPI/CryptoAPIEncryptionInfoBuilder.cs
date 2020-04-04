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
    using System.Diagnostics;
    using NPOI.POIFS.Crypt;
    using NPOI.Util;

    public class CryptoAPIEncryptionInfoBuilder : IEncryptionInfoBuilder
    {
        EncryptionInfo info;
        CryptoAPIEncryptionHeader header;
        CryptoAPIEncryptionVerifier verifier;
        CryptoAPIDecryptor decryptor;
        CryptoAPIEncryptor encryptor;

        public CryptoAPIEncryptionInfoBuilder()
        {
        }

        /**
         * Initialize the builder from a stream
         */
        public void Initialize(EncryptionInfo info, ILittleEndianInput dis)
        {
            this.info = info;
            int hSize = dis.ReadInt();
            header = new CryptoAPIEncryptionHeader(dis);
            verifier = new CryptoAPIEncryptionVerifier(dis, header);
            decryptor = new CryptoAPIDecryptor(this);
            encryptor = new CryptoAPIEncryptor(this);
        }

        /**
         * Initialize the builder from scratch
         */
        public void Initialize(EncryptionInfo info,
                CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm,
                int keyBits, int blockSize, ChainingMode chainingMode)
        {
            this.info = info;
            if (cipherAlgorithm == null) cipherAlgorithm = CipherAlgorithm.rc4;
            if (hashAlgorithm == null) hashAlgorithm = HashAlgorithm.sha1;
            if (keyBits == -1) keyBits = 0x28;
            Debug.Assert(cipherAlgorithm == CipherAlgorithm.rc4 && hashAlgorithm == HashAlgorithm.sha1);

            header = new CryptoAPIEncryptionHeader(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode);
            verifier = new CryptoAPIEncryptionVerifier(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode);
            decryptor = new CryptoAPIDecryptor(this);
            encryptor = new CryptoAPIEncryptor(this);
        }

        public CryptoAPIEncryptionHeader GetHeader()
        {
            return header;
        }

        public CryptoAPIEncryptionVerifier GetVerifier()
        {
            return verifier;
        }

        public CryptoAPIDecryptor GetDecryptor()
        {
            return decryptor;
        }

        public CryptoAPIEncryptor GetEncryptor()
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
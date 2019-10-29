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
    using NPOI.Util;

    public class BinaryRC4EncryptionInfoBuilder : IEncryptionInfoBuilder
    {

        EncryptionInfo info;
        BinaryRC4EncryptionHeader header;
        BinaryRC4EncryptionVerifier verifier;
        BinaryRC4Decryptor decryptor;
        BinaryRC4Encryptor encryptor;

        public BinaryRC4EncryptionInfoBuilder()
        {
        }

        public void Initialize(EncryptionInfo info, ILittleEndianInput dis)
        {
            this.info = info;
            int vMajor = info.VersionMajor;
            int vMinor = info.VersionMinor;
            Debug.Assert(vMajor == 1 && vMinor == 1);

            header = new BinaryRC4EncryptionHeader();
            verifier = new BinaryRC4EncryptionVerifier(dis);
            decryptor = new BinaryRC4Decryptor(this);
            encryptor = new BinaryRC4Encryptor(this);
        }

        public void Initialize(EncryptionInfo info,
            CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm,
            int keyBits, int blockSize, ChainingMode chainingMode)
        {
            this.info = info;
            header = new BinaryRC4EncryptionHeader();
            verifier = new BinaryRC4EncryptionVerifier();
            decryptor = new BinaryRC4Decryptor(this);
            encryptor = new BinaryRC4Encryptor(this);
        }

        public BinaryRC4EncryptionHeader GetHeader()
        {
            return header;
        }

        public BinaryRC4EncryptionVerifier GetVerifier()
        {
            return verifier;
        }

        public BinaryRC4Decryptor GetDecryptor()
        {
            return decryptor;
        }

        public BinaryRC4Encryptor GetEncryptor()
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
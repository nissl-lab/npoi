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

using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.POIFS.Crypt.XOR
{
    internal class XOREncryptionInfoBuilder : IEncryptionInfoBuilder
    {
        EncryptionInfo info;
        XOREncryptionHeader header;
        XOREncryptionVerifier verifier;
        XORDecryptor decryptor;
        XOREncryptor encryptor;

        public XOREncryptionInfoBuilder()
        {
        }

        public void Initialize(EncryptionInfo info, ILittleEndianInput dis)
        {
            this.info = info;
            header = new XOREncryptionHeader();
            verifier = new XOREncryptionVerifier(dis);
            decryptor = new XORDecryptor(this);
            encryptor = new XOREncryptor(this);
        }

        public void Initialize(EncryptionInfo info,
            CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm,
            int keyBits, int blockSize, ChainingMode chainingMode)
        {
            this.info = info;
            header = new XOREncryptionHeader();
            verifier = new XOREncryptionVerifier();
            decryptor = new XORDecryptor(this);
            encryptor = new XOREncryptor(this);
        }

        public XOREncryptionHeader GetHeader()
        {
            return header;
        }

        public XOREncryptionVerifier GetVerifier()
        {
            return verifier;
        }

        public XORDecryptor GetDecryptor()
        {
            return decryptor;
        }

        public XOREncryptor GetEncryptor()
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

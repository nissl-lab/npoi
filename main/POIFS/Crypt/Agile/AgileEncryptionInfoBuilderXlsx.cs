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

using NPOI.POIFS.FileSystem;
using NPOI.Util;

namespace NPOI.POIFS.Crypt.Agile
{
    public class AgileEncryptionInfoBuilderXlsx : IEncryptionInfoBuilder
    {
        private Encryptor encryptor;

        private EncryptionInfo info;

        public void Initialize(EncryptionInfo ei, ILittleEndianInput dis)
        {
            info = ei;
            encryptor = new AgileEncryptorForXlsx();
        }

        public void Initialize(EncryptionInfo ei, CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm,
            int keyBits, int blockSize, ChainingMode chainingMode)
        {
            info = ei;
            encryptor = new AgileEncryptorForXlsx();
        }

        EncryptionHeader IEncryptionInfoBuilder.GetHeader()
        {
            return null;
        }

        EncryptionVerifier IEncryptionInfoBuilder.GetVerifier()
        {
            return null;
        }

        Decryptor IEncryptionInfoBuilder.GetDecryptor()
        {
            return null;
        }

        Encryptor IEncryptionInfoBuilder.GetEncryptor()
        {
            return encryptor;
        }

        public Encryptor GetEncryptor()
        {
            return encryptor;
        }

        public EncryptionInfo GetInfo()
        {
            return info;
        }

        public static object ParseDescriptor(string descriptor)
        {
            return null;
        }

        protected static object ParseDescriptor(DocumentInputStream descriptor)
        {
            return null;
        }
    }
}
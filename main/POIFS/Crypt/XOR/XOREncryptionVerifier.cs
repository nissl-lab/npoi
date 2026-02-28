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

using NPOI.POIFS.Crypt.Standard;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.POIFS.Crypt.XOR
{
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.Standard;
    using NPOI.Util;

    public class XOREncryptionVerifier : EncryptionVerifier, EncryptionRecord
    {
        private int encryptedKey;
        private int encryptedVerifier;

        public XOREncryptionVerifier()
        {
        }

        public XOREncryptionVerifier(ILittleEndianInput is1)
        {
            encryptedKey = is1.ReadUShort();
            encryptedVerifier = is1.ReadUShort();
        }

        protected XOREncryptionVerifier(XOREncryptionVerifier other) 
        {
            encryptedKey = other.encryptedKey;
            encryptedVerifier = other.encryptedVerifier;
        }

        public int GetEncryptedKey()
        {
            return encryptedKey;
        }

        public void SetEncryptedKey(byte[] encryptedKey)
        {
            this.encryptedKey = LittleEndian.GetUShort(encryptedKey);
        }

        public int GetEncryptedVerifier()
        {
            return encryptedVerifier;
        }

        public void SetEncryptedVerifier(byte[] encryptedVerifier)
        {
            this.encryptedVerifier = LittleEndian.GetUShort(encryptedVerifier);
        }

        public void Write(LittleEndianByteArrayOutputStream bos)
        {
            bos.WriteShort(encryptedKey);
            bos.WriteShort(encryptedVerifier);
        }

        public XOREncryptionVerifier Copy()
        {
            return new XOREncryptionVerifier(this);
        }
    }
}

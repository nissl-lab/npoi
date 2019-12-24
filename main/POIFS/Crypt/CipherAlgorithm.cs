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

namespace NPOI.POIFS.Crypt
{
    using System;

    public class CipherAlgorithm
    {
        // key size for rc4: 0x00000028 - 0x00000080 (inclusive) with 8-bit increments
        // no block size, because its a streaming cipher
        public static CipherAlgorithm rc4 = new CipherAlgorithm(CipherProvider.rc4, "RC4", 0x6801, 0x40,
            new int[] { 0x28, 0x30, 0x38, 0x40, 0x48, 0x50, 0x58, 0x60, 0x68, 0x70, 0x78, 0x80 }, -1, 20, "RC4", false)
        { name = "rc4" };

        // aes has always a block size of 128 - only its keysize may vary
        public static CipherAlgorithm aes128 = new CipherAlgorithm(CipherProvider.aes, "AES", 0x660E, 128,
            new int[] {128}, 16, 32, "AES", false)
        { name = "aes128" };

        public static CipherAlgorithm aes192 = new CipherAlgorithm(CipherProvider.aes, "AES", 0x660F, 192,
            new int[] {192}, 16, 32, "AES", false)
        { name = "aes192" };

        public static CipherAlgorithm aes256 = new CipherAlgorithm(CipherProvider.aes, "AES", 0x6610, 256,
            new int[] {256}, 16, 32, "AES", false)
        { name = "aes256" };

        public static CipherAlgorithm rc2 = new CipherAlgorithm(null, "RC2", -1, 0x80,
            new int[] {0x28, 0x30, 0x38, 0x40, 0x48, 0x50, 0x58, 0x60, 0x68, 0x70, 0x78, 0x80}, 8, 20, "RC2", false)
        { name = "rc2" };

        public static CipherAlgorithm des = new CipherAlgorithm(null, "DES", -1, 64, new int[] {64}, 8 /*for 56-bit*/,
            32, "DES", false)
        { name = "des" };

        // desx is not supported. Not sure, if it can be simulated by des3 somehow
        public static CipherAlgorithm des3 = new CipherAlgorithm(null, "DESede", -1, 192, new int[] {192}, 8, 32, "3DES",
            false)
        { name = "des3" };

        // need bouncycastle provider for this one ...
        // see http://stackoverflow.com/questions/4436397/3des-des-encryption-using-the-jce-generating-an-acceptable-key
        public static CipherAlgorithm des3_112 = new CipherAlgorithm(null, "DESede", -1, 128, new int[] {128}, 8, 32,
            "3DES_112", true)
        { name = "des3_112" };

        // only for digital signatures
        public static CipherAlgorithm rsa = new CipherAlgorithm(null, "RSA", -1, 1024,
            new int[] {1024, 2048, 3072, 4096}, -1, -1, "", false)
        { name = "rsa" };

        public static CipherAlgorithm[] Values = new CipherAlgorithm[] { rc4, aes128, aes192, aes256, rc2, des, des3, des3_112, rsa };
        public CipherProvider provider;
        public String jceId;
        public int ecmaId;
        public int defaultKeySize;
        public int[] allowedKeySize;
        public int blockSize;
        public int encryptedVerifierHashLength;
        public String xmlId;
        public bool needsBouncyCastle;
        private string name;
        
        public static CipherAlgorithm ValueOf(string alg)
        {
            switch (alg.ToLower())
            {
                case "rc4": return rc4;
                case "rc2": return rc2;
                case "aes128": return aes128;
                case "aes192": return aes192;
                case "aes256": return aes256;
                case "des": return des;
                case "des3": return des3;
                case "des3_112": return des3_112;
                case "rsa": return rsa;
                default:
                    throw new ArgumentException(string.Format("not found definition of cipher algorithm {0}", alg));
            }
        }

        public override string ToString()
        {
            return name;
        }

        public CipherAlgorithm(CipherProvider provider, String jceId, int ecmaId, int defaultKeySize,
            int[] allowedKeySize, int blockSize, int encryptedVerifierHashLength, String xmlId, bool needsBouncyCastle)
        {
            this.provider = provider;
            this.jceId = jceId;
            this.ecmaId = ecmaId;
            this.defaultKeySize = defaultKeySize;
            this.allowedKeySize = (int[])allowedKeySize.Clone();
            this.blockSize = blockSize;
            this.encryptedVerifierHashLength = encryptedVerifierHashLength;
            this.xmlId = xmlId;
            this.needsBouncyCastle = needsBouncyCastle;
        }

        public static CipherAlgorithm FromEcmaId(int ecmaId)
        {
            foreach (CipherAlgorithm ca in CipherAlgorithm.Values)
            {
                if (ca.ecmaId == ecmaId) return ca;
            }
            throw new EncryptedDocumentException("cipher algorithm " + ecmaId + " not found");
        }

        public static CipherAlgorithm FromXmlId(String xmlId, int keySize)
        {
            foreach (CipherAlgorithm ca in CipherAlgorithm.Values)
            {
                if (!ca.xmlId.Equals(xmlId)) continue;
                foreach (int ks in ca.allowedKeySize)
                {
                    if (ks == keySize) return ca;
                }
            }
            throw new EncryptedDocumentException("cipher algorithm " + xmlId + "/" + keySize + " not found");
        }
    }
}
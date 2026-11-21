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

    public class HashAlgorithm
    {
        public static readonly HashAlgorithm none = new HashAlgorithm("", 0x0000, "", 0, "", false);
        public static readonly HashAlgorithm sha1 = new HashAlgorithm("SHA-1", 0x8004, "SHA1", 20, "HmacSHA1", false);

        public static readonly HashAlgorithm sha256 = new HashAlgorithm("SHA-256", 0x800C, "SHA256", 32, "HmacSHA256",
            false);

        public static readonly HashAlgorithm sha384 = new HashAlgorithm("SHA-384", 0x800D, "SHA384", 48, "HmacSHA384",
            false);

        public static readonly HashAlgorithm sha512 = new HashAlgorithm("SHA-512", 0x800E, "SHA512", 64, "HmacSHA512",
            false);

        /* only for agile encryption */
        public static readonly HashAlgorithm md5 = new HashAlgorithm("MD5", -1, "MD5", 16, "HmacMD5", false);
        // although sunjc2 supports md2, hmac-md2 is only supported by bouncycastle
        public static readonly HashAlgorithm md2 = new HashAlgorithm("MD2", -1, "MD2", 16, "Hmac-MD2", true);
        public static readonly HashAlgorithm md4 = new HashAlgorithm("MD4", -1, "MD4", 16, "Hmac-MD4", true);

        public static readonly HashAlgorithm ripemd128 = new HashAlgorithm("RipeMD128", -1, "RIPEMD-128", 16,
            "HMac-RipeMD128", true);

        public static readonly HashAlgorithm ripemd160 = new HashAlgorithm("RipeMD160", -1, "RIPEMD-160", 20,
            "HMac-RipeMD160", true);

        public static readonly HashAlgorithm whirlpool = new HashAlgorithm("Whirlpool", -1, "WHIRLPOOL", 64,
            "HMac-Whirlpool", true);

        // only for xml signing
        public static readonly HashAlgorithm sha224 = new HashAlgorithm("SHA-224", -1, "SHA224", 28, "HmacSHA224", true);

        public static readonly HashAlgorithm rc4 = new HashAlgorithm("SHA-1", 32772, "SHA1", 128, "HmacSHA1", false);

        public static HashAlgorithm[] values = new HashAlgorithm[] {none,sha1,sha256,sha384,sha512,md5,md4,md2,ripemd128,ripemd160,whirlpool,sha224,rc4};

        public String jceId;
        public int ecmaId;
        public String ecmaString;
        public int hashSize;
        public String jceHmacId;
        public bool needsBouncyCastle;

        public HashAlgorithm(String jceId, int ecmaId, String ecmaString, int hashSize, String jceHmacId,
            bool needsBouncyCastle)
        {
            this.jceId = jceId;
            this.ecmaId = ecmaId;
            this.ecmaString = ecmaString;
            this.hashSize = hashSize;
            this.jceHmacId = jceHmacId;
            this.needsBouncyCastle = needsBouncyCastle;
        }

        public static HashAlgorithm FromEcmaId(int ecmaId)
        {
            foreach (HashAlgorithm ha in values)
            {
                if (ha.ecmaId == ecmaId) return ha;
            }
            throw new EncryptedDocumentException("hash algorithm not found");
        }

        public static HashAlgorithm FromEcmaId(String ecmaString)
        {
            foreach (HashAlgorithm ha in values)
            {
                if (ha.ecmaString.Equals(ecmaString)) return ha;
            }
            throw new EncryptedDocumentException("hash algorithm not found");
        }

        public static HashAlgorithm FromString(String string1)
        {
            foreach (HashAlgorithm ha in values)
            {
                if (ha.ecmaString.Equals(string1, StringComparison.CurrentCultureIgnoreCase) || ha.jceId.Equals(string1, StringComparison.CurrentCultureIgnoreCase)) return ha;
            }
            throw new EncryptedDocumentException("hash algorithm not found");
        }
    }
}
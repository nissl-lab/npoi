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


/**
 * Reads and Processes OOXML Encryption Headers
 * The constants are largely based on ZIP constants.
 */

    public abstract class EncryptionHeader
    {
        public static int ALGORITHM_RC4 = CipherAlgorithm.rc4.ecmaId;
        public static int ALGORITHM_AES_128 = CipherAlgorithm.aes128.ecmaId;
        public static int ALGORITHM_AES_192 = CipherAlgorithm.aes192.ecmaId;
        public static int ALGORITHM_AES_256 = CipherAlgorithm.aes256.ecmaId;

        public static int HASH_NONE = HashAlgorithm.none.ecmaId;
        public static int HASH_SHA1 = HashAlgorithm.sha1.ecmaId;
        public static int HASH_SHA256 = HashAlgorithm.sha256.ecmaId;
        public static int HASH_SHA384 = HashAlgorithm.sha384.ecmaId;
        public static int HASH_SHA512 = HashAlgorithm.sha512.ecmaId;

        public static int PROVIDER_RC4 = CipherProvider.rc4.ecmaId;
        public static int PROVIDER_AES = CipherProvider.aes.ecmaId;

        public static int MODE_ECB = ChainingMode.ecb.ecmaId;
        public static int MODE_CBC = ChainingMode.cbc.ecmaId;
        public static int MODE_CFB = ChainingMode.cfb.ecmaId;

        protected EncryptionHeader()
        {
        }

        /**
     * @deprecated use GetChainingMode().ecmaId
     */
        [Obsolete("use ChainingMode.ecmaId", true)]
        public int GetCipherMode()
        {
            return ChainingMode.ecmaId;
        }

        public ChainingMode ChainingMode { get; set; }


        public int Flags { get; set; }

        public int SizeExtra { get; set; }

        [Obsolete("use CipherAlgorithm")]
        public int GetAlgorithm()
        {
            return CipherAlgorithm.ecmaId;
        }

        public CipherAlgorithm CipherAlgorithm { get; set; }

        [Obsolete("use HashAlgorithmEx")]
        public int GetHashAlgorithm()
        {
            return HashAlgorithm.ecmaId;
        }

        public HashAlgorithm HashAlgorithm { get; set; }

        public int KeySize { get; set; }

        public int BlockSize { get; set; }

        public byte[] KeySalt { get; set; }

        [Obsolete("use CipherProvider")]
        public int GetProviderType()
        {
            return CipherProvider.ecmaId;
        }

        public CipherProvider CipherProvider { get; set; }

        public string CspName { get; set; }
    }

}
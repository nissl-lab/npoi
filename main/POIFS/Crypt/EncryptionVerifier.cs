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
 * Used when Checking if a key is valid for a document 
 */

    public abstract class EncryptionVerifier
    {
        private byte[] encryptedVerifier;
        private byte[] encryptedVerifierHash;
        private byte[] encryptedKey;
        // protected int verifierHashSize;
        private int spinCount;
        private CipherAlgorithm cipherAlgorithm;
        private ChainingMode chainingMode;
        private HashAlgorithm hashAlgorithm;

        protected EncryptionVerifier()
        {
        }

        public byte[] Salt { get; set; }
        /**
     * The method name is misleading - you'll Get the encrypted verifier, not the plain verifier
     * @deprecated use GetEncryptedVerifier()
     */

        public byte[] GetVerifier()
        {
            return encryptedVerifier;
        }

        public byte[] EncryptedVerifier { get; set; }
        

        /**
     * The method name is misleading - you'll Get the encrypted verifier hash, not the plain verifier hash
     * @deprecated use GetEnryptedVerifierHash
     */

        public byte[] GetVerifierHash()
        {
            return encryptedVerifierHash;
        }

        public byte[] EncryptedVerifierHash { get; set; }

        public int SpinCount { get; set; }

        public int GetCipherMode()
        {
            return chainingMode.ecmaId;
        }

        public int GetAlgorithm()
        {
            return cipherAlgorithm.ecmaId;
        }

        /**
     * @deprecated use GetCipherAlgorithm().jceId
     */

        public String GetAlgorithmName()
        {
            return cipherAlgorithm.jceId;
        }

        public byte[] EncryptedKey { get; set; }

        public CipherAlgorithm CipherAlgorithm { get; set; }

        public HashAlgorithm HashAlgorithm { get; set; }

        public ChainingMode ChainingMode { get; set; }        
    }

}
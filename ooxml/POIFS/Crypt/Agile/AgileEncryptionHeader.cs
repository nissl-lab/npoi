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
namespace NPOI.POIFS.Crypt.Agile
{
    using System;
    using NPOI.OpenXmlFormats.Encryption;
    using NPOI.POIFS.Crypt;


    public class AgileEncryptionHeader : EncryptionHeader
    {
        private byte[] encryptedHmacKey, encryptedHmacValue;

        public AgileEncryptionHeader(string descriptor)
                : this(AgileEncryptionInfoBuilder.ParseDescriptor(descriptor))
        {

        }

        protected internal AgileEncryptionHeader(EncryptionDocument ed)
        {
            CT_KeyData keyData;
            try
            {
                keyData = ed.GetEncryption().keyData;
                if (keyData == null)
                {
                    throw new NullReferenceException("keyData not Set");
                }
            }
            catch
            {
                throw new EncryptedDocumentException("Unable to parse keyData");
            }

            KeySize = ((int)keyData.keyBits);
            Flags = (0);
            SizeExtra = (0);
            CspName = (null);
            BlockSize = (int)(keyData.blockSize);

            int keyBits = (int)keyData.keyBits;

            CipherAlgorithm ca = CipherAlgorithm.FromXmlId(keyData.cipherAlgorithm.ToString(), keyBits);
            CipherAlgorithm = (ca);
            CipherProvider = (ca.provider);

            switch (keyData.cipherChaining)
            {
                case ST_CipherChaining.ChainingModeCBC:
                    ChainingMode = (ChainingMode.cbc);
                    break;
                case ST_CipherChaining.ChainingModeCFB:
                    ChainingMode = (ChainingMode.cfb);
                    break;
                default:
                    throw new EncryptedDocumentException("Unsupported chaining mode - " + keyData.cipherChaining.ToString());
            }

            int hashSize = (int)keyData.hashSize;

            HashAlgorithm ha = HashAlgorithm.FromEcmaId(keyData.hashAlgorithm.ToString());
            HashAlgorithm = (ha);

            if (HashAlgorithm.hashSize != hashSize)
            {
                throw new EncryptedDocumentException("Unsupported hash algorithm: " +
                        keyData.hashAlgorithm + " @ " + hashSize + " bytes");
            }

            int saltLength = (int)keyData.saltSize;
            SetKeySalt(keyData.saltValue);
            if (KeySalt.Length != saltLength)
            {
                throw new EncryptedDocumentException("Invalid salt length");
            }

            CT_DataIntegrity di = ed.GetEncryption().dataIntegrity;
            SetEncryptedHmacKey(di.encryptedHmacKey);
            SetEncryptedHmacValue(di.encryptedHmacValue);
        }


        public AgileEncryptionHeader(CipherAlgorithm algorithm, HashAlgorithm hashAlgorithm, int keyBits, int blockSize, ChainingMode chainingMode)
        {
            CipherAlgorithm = (algorithm);
            HashAlgorithm = (hashAlgorithm);
            KeySize = (keyBits);
            BlockSize = (blockSize);
            ChainingMode = (chainingMode);
        }

        // make method visible for this package
        protected void SetKeySalt(byte[] salt)
        {
            if (salt == null || salt.Length != BlockSize)
            {
                throw new EncryptedDocumentException("invalid verifier salt");
            }
            base.KeySalt = (salt);
        }

        public byte[] GetEncryptedHmacKey()
        {
            return encryptedHmacKey;
        }

        public void SetEncryptedHmacKey(byte[] encryptedHmacKey)
        {
            this.encryptedHmacKey = encryptedHmacKey;
        }

        public byte[] GetEncryptedHmacValue()
        {
            return encryptedHmacValue;
        }

        public void SetEncryptedHmacValue(byte[] encryptedHmacValue)
        {
            this.encryptedHmacValue = encryptedHmacValue;
        }
    }

}
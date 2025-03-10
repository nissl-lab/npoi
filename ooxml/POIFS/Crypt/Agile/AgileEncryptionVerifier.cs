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
    using System.Collections.Generic;
    using System.IO;
    using NPOI.OpenXmlFormats.Encryption;
    using NPOI.POIFS.Crypt;
    using Org.BouncyCastle.Asn1.X509;
    using Org.BouncyCastle.X509;

    /**
     * Used when Checking if a key is valid for a document 
     */
    public class AgileEncryptionVerifier : EncryptionVerifier
    {

        public class AgileCertificateEntry
        {
            internal X509Certificate x509;
            internal byte[] encryptedKey;
            internal byte[] certVerifier;
        }

        private List<AgileCertificateEntry> certList = new List<AgileCertificateEntry>();

        public AgileEncryptionVerifier(String descriptor)
            : this(AgileEncryptionInfoBuilder.ParseDescriptor(descriptor))
        {
            
        }

        protected internal AgileEncryptionVerifier(EncryptionDocument ed)
        {
            IEnumerator<CT_KeyEncryptor> encList = ed.GetEncryption().keyEncryptors.keyEncryptor.GetEnumerator();
            CT_PasswordKeyEncryptor keyData;
            try
            {
                //keyData = encList.Next().EncryptedPasswordKey;
                encList.MoveNext();
                keyData = encList.Current.Item as CT_PasswordKeyEncryptor;
                if (keyData == null)
                {
                    throw new NullReferenceException("encryptedKey not Set");
                }
            }
            catch (Exception e)
            {
                throw new EncryptedDocumentException("Unable to parse keyData", e);
            }

            int keyBits = (int)keyData.keyBits;

            CipherAlgorithm ca = CipherAlgorithm.FromXmlId(keyData.cipherAlgorithm.ToString(), keyBits);
            CipherAlgorithm = (ca);

            int hashSize = (int)keyData.hashSize;

            HashAlgorithm ha = HashAlgorithm.FromEcmaId(keyData.hashAlgorithm.ToString());
            HashAlgorithm = (ha);

            if (HashAlgorithm.hashSize != hashSize)
            {
                throw new EncryptedDocumentException("Unsupported hash algorithm: " +
                        keyData.hashAlgorithm + " @ " + hashSize + " bytes");
            }

            SpinCount = (int)(keyData.spinCount);
            EncryptedVerifier = (keyData.encryptedVerifierHashInput);
            Salt = (keyData.saltValue);
            EncryptedKey = (keyData.encryptedKeyValue);
            EncryptedVerifierHash = (keyData.encryptedVerifierHashValue);

            int saltSize = (int)keyData.saltSize;
            if (saltSize != Salt.Length)
                throw new EncryptedDocumentException("Invalid salt size");

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
            //if (!encList.HasNext()) return;

            try
            {
                //CertificateFactory cf = CertificateFactory.GetInstance("X.509");
                while (encList.MoveNext())
                {
                    CT_CertificateKeyEncryptor certKey = encList.Current.Item as CT_CertificateKeyEncryptor;
                    AgileCertificateEntry ace = new AgileCertificateEntry();
                    ace.certVerifier = certKey.certVerifier;
                    ace.encryptedKey = certKey.encryptedKeyValue;
                    ace.x509 = new X509Certificate(X509CertificateStructure.GetInstance(certKey.X509Certificate));
                    certList.Add(ace);
                }
            }
            catch (Exception e)
            {
                throw new EncryptedDocumentException("can't parse X509 certificate", e);
            }
        }

        public AgileEncryptionVerifier(CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm, int keyBits, int blockSize, ChainingMode chainingMode)
        {
            CipherAlgorithm = (cipherAlgorithm);
            HashAlgorithm = (hashAlgorithm);
            ChainingMode = (chainingMode);
            SpinCount = (100000); // TODO: use parameter
        }

        protected void SetSalt(byte[] salt)
        {
            if (salt == null || salt.Length != CipherAlgorithm.blockSize)
            {
                throw new EncryptedDocumentException("invalid verifier salt");
            }
            base.Salt = (/*setter*/salt);
        }

        // make method visible for this package
        protected void SetEncryptedVerifier(byte[] encryptedVerifier)
        {
            base.EncryptedVerifier = (/*setter*/encryptedVerifier);
        }

        // make method visible for this package
        protected void SetEncryptedVerifierHash(byte[] encryptedVerifierHash)
        {
            base.EncryptedVerifierHash = (/*setter*/encryptedVerifierHash);
        }

        // make method visible for this package
        protected void SetEncryptedKey(byte[] encryptedKey)
        {
            base.EncryptedKey = (/*setter*/encryptedKey);
        }

        public void AddCertificate(X509Certificate x509)
        {
            AgileCertificateEntry ace = new AgileCertificateEntry();
            ace.x509 = x509;
            certList.Add(ace);
        }

        public List<AgileCertificateEntry> GetCertificates()
        {
            return certList;
        }
    }

    internal sealed class CertificateFactory
    {
        internal static CertificateFactory GetInstance(string v)
        {
            throw new NotImplementedException();
        }

        internal X509Certificate GenerateCertificate(MemoryStream memoryStream)
        {
            throw new NotImplementedException();
        }
    }
}

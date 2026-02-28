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
    using System.IO;
    using System.Text;
    using System.Xml;
    using NPOI.OpenXmlFormats.Encryption;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public class AgileEncryptionInfoBuilder : IEncryptionInfoBuilder
    {

        EncryptionInfo info;
        AgileEncryptionHeader header;
        AgileEncryptionVerifier verifier;
        AgileDecryptor decryptor;
        AgileEncryptor encryptor;

        public void Initialize(EncryptionInfo info, ILittleEndianInput dis)
        {
            this.info = info;

            EncryptionDocument ed = ParseDescriptor((DocumentInputStream)dis);
            header = new AgileEncryptionHeader(ed);
            verifier = new AgileEncryptionVerifier(ed);
            if (info.VersionMajor == EncryptionMode.Agile.VersionMajor
                && info.VersionMinor == EncryptionMode.Agile.VersionMinor)
            {
                decryptor = new AgileDecryptor(this);
                encryptor = new AgileEncryptor(this);
            }
        }

        public void Initialize(EncryptionInfo info, CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm, int keyBits, int blockSize, ChainingMode chainingMode)
        {
            this.info = info;

            if (cipherAlgorithm == null)
            {
                cipherAlgorithm = CipherAlgorithm.aes128;
            }
            if (cipherAlgorithm == CipherAlgorithm.rc4)
            {
                throw new EncryptedDocumentException("RC4 must not be used with agile encryption.");
            }
            if (hashAlgorithm == null)
            {
                hashAlgorithm = HashAlgorithm.sha1;
            }
            if (chainingMode == null)
            {
                chainingMode = ChainingMode.cbc;
            }
            if (!(chainingMode == ChainingMode.cbc || chainingMode == ChainingMode.cfb))
            {
                throw new EncryptedDocumentException("Agile encryption only supports CBC/CFB chaining.");
            }
            if (keyBits == -1)
            {
                keyBits = cipherAlgorithm.defaultKeySize;
            }
            if (blockSize == -1)
            {
                blockSize = cipherAlgorithm.blockSize;
            }
            bool found = false;
            foreach (int ks in cipherAlgorithm.allowedKeySize)
            {
                found |= (ks == keyBits);
            }
            if (!found)
            {
                throw new EncryptedDocumentException("KeySize " + keyBits + " not allowed for Cipher " + cipherAlgorithm.ToString());
            }
            header = new AgileEncryptionHeader(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode);
            verifier = new AgileEncryptionVerifier(cipherAlgorithm, hashAlgorithm, keyBits, blockSize, chainingMode);
            decryptor = new AgileDecryptor(this);
            encryptor = new AgileEncryptor(this);
        }

        public AgileEncryptionHeader GetHeader()
        {
            return header;
        }

        public AgileEncryptionVerifier GetVerifier()
        {
            return verifier;
        }

        public AgileDecryptor GetDecryptor()
        {
            return decryptor;
        }

        public AgileEncryptor GetEncryptor()
        {
            return encryptor;
        }

        public EncryptionInfo GetEncryptionInfo()
        {
            return info;
        }

        public static EncryptionDocument ParseDescriptor(String descriptor)
        {
            try
            {
                return EncryptionDocument.Parse(descriptor);
            }
            catch (XmlException e)
            {
                throw new EncryptedDocumentException("Unable to parse encryption descriptor", e);
            }
        }

        protected static EncryptionDocument ParseDescriptor(DocumentInputStream descriptor)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                byte[] buf = new byte[descriptor.Length - descriptor.Position];
                descriptor.ReadFully(buf);
                string xml = Encoding.UTF8.GetString(buf);
                OpenXml4Net.Util.XmlHelper.LoadXmlSafe(xmlDoc, xml, Encoding.UTF8);

                return EncryptionDocument.Parse(xmlDoc);
            }
            catch (Exception e)
            {
                throw new EncryptedDocumentException("Unable to parse encryption descriptor", e);
            }
        }

        EncryptionHeader IEncryptionInfoBuilder.GetHeader()
        {
            return this.header;
        }

        EncryptionVerifier IEncryptionInfoBuilder.GetVerifier()
        {
            return this.verifier;
        }

        Decryptor IEncryptionInfoBuilder.GetDecryptor()
        {
            return this.decryptor;
        }

        Encryptor IEncryptionInfoBuilder.GetEncryptor()
        {
            return this.encryptor;
        }
    }
}

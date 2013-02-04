/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Text;
using System.Xml;
using System.IO;
using NPOI.POIFS.FileSystem;

namespace NPOI.POIFS.Crypt
{
    public class EncryptionVerifier
    {
        private byte[] salt;
        private byte[] verifier;
        private byte[] verifierHash;
        private byte[] encryptedKey;
        private int verifierHashSize;
        private int spinCount;
        private int algorithm;
        private int cipherMode;

        public EncryptionVerifier(string descriptor)
        {
            XmlAttributeCollection keyData = null;
            try
            {
                MemoryStream ms = new MemoryStream(Encoding.Default.GetBytes(descriptor));
                XmlDocument xml = new XmlDocument();
                xml.Load(ms);
                XmlNodeList nodes = xml.GetElementsByTagName("keyEncryptor");
                XmlNodeList keyEncryptor = nodes[0].ChildNodes;

                for (int i = 0; i < keyEncryptor.Count; i++)
                {
                    XmlNode node = keyEncryptor[i];
                    if (node.Name.Equals("p:encryptedKey"))
                    {
                        keyData = node.Attributes;
                        break;
                    }
                }

                if (keyData == null)
                    throw new EncryptedDocumentException("");

                spinCount = Int32.Parse(keyData.GetNamedItem("spinCount").Value);
                verifier = Convert.FromBase64String(keyData.GetNamedItem("encryptedVerifierHashInput").Value);
                salt = Convert.FromBase64String(keyData.GetNamedItem("saltValue").Value);
                encryptedKey = Convert.FromBase64String(keyData.GetNamedItem("encryptedKeyValue").Value);

                int saltSize = Int32.Parse(keyData.GetNamedItem("saltSize").Value);

                if (saltSize != salt.Length)
                    throw new EncryptedDocumentException("Invalid salt size");

                verifierHash = Convert.FromBase64String(keyData.GetNamedItem("encryptedVerifierHashValue").Value);

                int blockSize = Int32.Parse(keyData.GetNamedItem("blockSize").Value);

                string alg = keyData.GetNamedItem("cipherAlgorithm").Value;

                if ("AES".Equals(alg))
                {
                    if (blockSize == 16)
                        algorithm = EncryptionHeader.ALGORITHM_AES_128;
                    else if (blockSize == 24)
                        algorithm = EncryptionHeader.ALGORITHM_AES_192;
                    else if (blockSize == 32)
                        algorithm = EncryptionHeader.ALGORITHM_AES_256;
                    else
                        throw new EncryptedDocumentException("Unsupported block size");
                }
                else
                    throw new EncryptedDocumentException("Unsupported cipher");

                string chain = keyData.GetNamedItem("cipherChaining").Value;

                if ("ChainingModeCBC".Equals(chain))
                    cipherMode = EncryptionHeader.MODE_CBC;
                else if ("ChainingModeCFB".Equals(chain))
                    cipherMode = EncryptionHeader.MODE_CFB;
                else
                    throw new EncryptedDocumentException("Unsupported chaining mode");

                verifierHashSize = Int32.Parse(keyData.GetNamedItem("hashSize").Value);

            }
            catch
            {
                throw new EncryptedDocumentException("Unable to parse keyEncryptor");
            }
        }

        public EncryptionVerifier(DocumentInputStream dis, int encryptedLength)
        {
            int saltSize = dis.ReadInt();

            if (saltSize != 16)
                throw new Exception("Salt size != 16 !?");

            salt = new byte[16];
            dis.ReadFully(salt);
            verifier = new byte[16];
            dis.ReadFully(verifier);

            verifierHashSize = dis.ReadInt();

            verifierHash = new byte[encryptedLength];
            dis.ReadFully(verifierHash);

            spinCount = 50000;
            algorithm = EncryptionHeader.ALGORITHM_AES_128;
            cipherMode = EncryptionHeader.MODE_ECB;
            encryptedKey = null;
        }

        public byte[] Salt
        {
            get { return salt; }
        }

        public byte[] Verifier
        {
            get { return verifier; }
        }
        public byte[] VerifierHash
        {
            get
            {
                return verifierHash;
            }
        }
        public int SpinCount
        {
            get { return spinCount; }
        }

        public int CipherMode
        {
            get { return cipherMode; }
        }

        public int Algorithm
        {
            get { return algorithm; }
        }

        public byte[] EncryptedKey
        {
            get { return encryptedKey; }
        }
    }
}



































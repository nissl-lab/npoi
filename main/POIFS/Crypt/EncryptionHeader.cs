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
using System.IO;
using System.Text;
using System.Xml;

using NPOI;
using NPOI.POIFS.FileSystem;


namespace NPOI.POIFS.Crypt
{
    public class EncryptionHeader
    {
        public const  int ALGORITHM_RC4 = 0x6801;
        public const int ALGORITHM_AES_128 = 0x660E;
        public const int ALGORITHM_AES_192 = 0x660F;
        public const int ALGORITHM_AES_256 = 0x6610;

        public const int HASH_SHA1 = 0x8004;

        public const int PROVIDER_RC4 = 1;
        public const int PROVIDER_AES = 0x18;

        public const int MODE_ECB = 1;
        public const int MODE_CBC = 2;
        public const int MODE_CFB = 3;

        private  int flags;
        private  int sizeExtra;
        private int algorithm;
        private int hashAlgorithm;
        private int keySize;
        private int providerType;
        private int cipherMode;
        private byte[] keySalt;
        private String cspName;

        public EncryptionHeader(DocumentInputStream dr)
        {
            flags = dr.ReadInt();
            sizeExtra = dr.ReadInt();
            algorithm = dr.ReadInt();
            hashAlgorithm = dr.ReadInt();
            keySize = dr.ReadInt();
            providerType = dr.ReadInt();

            dr.ReadLong();  //skip reserved.

            StringBuilder builder = new StringBuilder();

            while (true)
            {
                char c = (char)dr.ReadShort();

                if (c == 0)
                    break;
                builder.Append(c);
            }

            cspName = builder.ToString();
            cipherMode = MODE_ECB;
            keySalt = null;
        }

        public EncryptionHeader(string descriptor)
        {
            try
            {
                XmlAttributeCollection keyData;
                try
                {
                    MemoryStream ms = new MemoryStream(Encoding.Default.GetBytes(descriptor));
                    XmlDocument xml = new XmlDocument();
                    xml.Load(ms);
                    XmlNodeList node = xml.GetElementsByTagName("keyData");
                    keyData = node[0].Attributes;
                }
                catch (Exception)
                {
                    throw new EncryptedDocumentException("Unable to parse keyData");
                }

                keySize = Int32.Parse(keyData.GetNamedItem("keyBits").Value);

                flags = 0;
                sizeExtra = 0;
                cspName = null;

                int blockSize = Int32.Parse(keyData.GetNamedItem("blockSize").Value);
                string cipher = keyData.GetNamedItem("cipherAlgorithm").Value;

                if ("AES".Equals(cipher))
                {
                    providerType = PROVIDER_AES;
                    if (blockSize == 16)
                        algorithm = ALGORITHM_AES_128;
                    else if (blockSize == 24)
                        algorithm = ALGORITHM_AES_192;
                    else if (blockSize == 32)
                        algorithm = ALGORITHM_AES_256;
                    else
                        throw new EncryptedDocumentException("Unsupported key length");
                }
                else
                {
                    throw new EncryptedDocumentException("Unsupported cipher");
                }

                string chaining = keyData.GetNamedItem("cipherChaining").Value;

                if ("ChainingModeCBC".Equals(chaining))
                    cipherMode = MODE_CBC;
                else if ("ChainingModeCFB".Equals(chaining))
                    cipherMode = MODE_CFB;
                else
                    throw new EncryptedDocumentException("Unsupported chaining mode");

                string hasAlg = keyData.GetNamedItem("hashAlgorithm").Value;
                int hashSize = Int32.Parse(keyData.GetNamedItem("hashSize").Value);

                if("SHA1".Equals(hasAlg) && hashSize == 20)
                    hashAlgorithm = HASH_SHA1;
                else
                    throw new EncryptedDocumentException("Unsupported hash algorithm");

                string salt = keyData.GetNamedItem("saltValue").Value;
                int saltLength = Int32.Parse(keyData.GetNamedItem("saltSize").Value);

                //keySalt = Base64.DecodeBase64(Encoding.Default.GetBytes(salt));
                //keySalt = Base64.DecodeBase64(salt);
                keySalt = Convert.FromBase64String(salt);
                if(keySalt.Length != saltLength)
                    throw new EncryptedDocumentException("Invalid salt length");


            }
            catch (IOException ex)
            {
                throw ex;
            }
        }

        public int CipherMode
        {
            get { return cipherMode; }
        }

        public int Flags
        {
            get { return flags; }
        }

        public int SizeExtra
        {
            get { return sizeExtra; }
        }

        public int Algorithm
        {
            get { return algorithm; }
        }

        public int HashAlgorithm
        {
            get { return hashAlgorithm; }
        }

        public int KeySize
        {
            get { return keySize; }
        }

        public byte[] KeySalt
        {
            get { return keySalt; }
        }

        public int ProviderType
        {
            get { return providerType; }
        }

        public string CspName
        {
            get { return cspName; }
        }

    }
}

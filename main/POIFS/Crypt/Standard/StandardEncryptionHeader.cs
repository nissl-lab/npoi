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
namespace NPOI.POIFS.Crypt.Standard
{
    using System;
    using NPOI.Util;
    using System.Text;

    public class StandardEncryptionHeader : EncryptionHeader, EncryptionRecord
    {

        internal StandardEncryptionHeader(ILittleEndianInput is1)
        {
            Flags = (is1.ReadInt());
            SizeExtra = (is1.ReadInt());
            CipherAlgorithm = (CipherAlgorithm.FromEcmaId(is1.ReadInt()));
            HashAlgorithm = (HashAlgorithm.FromEcmaId(is1.ReadInt()));
            int keySize = is1.ReadInt();
            if (keySize == 0)
            {
                // for the sake of inheritance of the cryptoAPI classes
                // see 2.3.5.1 RC4 CryptoAPI Encryption Header
                // If Set to 0x00000000, it MUST be interpreted as 0x00000028 bits.
                keySize = 0x28;
            }
            KeySize = (keySize);
            BlockSize = (keySize);
            CipherProvider = (CipherProvider.FromEcmaId(is1.ReadInt()));

            is1.ReadLong(); // skip reserved

            // CSPName may not always be specified
            // In some cases, the salt value of the EncryptionVerifier is the next chunk of data
            ((ByteArrayInputStream)is1).Mark(LittleEndianConsts.INT_SIZE + 1);
            int CheckForSalt = is1.ReadInt();
            ((ByteArrayInputStream)is1).Reset();

            if (CheckForSalt == 16)
            {
                CspName = ("");
            }
            else
            {
                StringBuilder builder = new StringBuilder();
                while (true)
                {
                    char c = (char)is1.ReadShort();
                    if (c == 0) break;
                    builder.Append(c);
                }
                CspName = (builder.ToString());
            }

            ChainingMode = (ChainingMode.ecb);
            KeySalt = (null);
        }

        protected internal StandardEncryptionHeader(CipherAlgorithm cipherAlgorithm, HashAlgorithm hashAlgorithm, int keyBits, int blockSize, ChainingMode chainingMode)
        {
            CipherAlgorithm = (cipherAlgorithm);
            HashAlgorithm = (hashAlgorithm);
            KeySize = (keyBits);
            BlockSize = (blockSize);
            CipherProvider = (cipherAlgorithm.provider);
            Flags = (EncryptionInfo.flagCryptoAPI.SetBoolean(0, true)
                     | EncryptionInfo.flagAES.SetBoolean(0, cipherAlgorithm.provider == CipherProvider.aes));
            // see http://msdn.microsoft.com/en-us/library/windows/desktop/bb931357(v=vs.85).aspx for a full list
            // SetCspName("Microsoft Enhanced RSA and AES Cryptographic Provider");
        }

        /**
         * Serializes the header 
         */
        public void Write(LittleEndianByteArrayOutputStream bos)
        {
            int startIdx = bos.WriteIndex;
            ILittleEndianOutput sizeOutput = bos.CreateDelayedOutput(LittleEndianConsts.INT_SIZE);
            bos.WriteInt(Flags);
            bos.WriteInt(0); // size extra
            bos.WriteInt(CipherAlgorithm.ecmaId);
            bos.WriteInt(HashAlgorithm.ecmaId);
            bos.WriteInt(KeySize);
            bos.WriteInt(CipherProvider.ecmaId);
            bos.WriteInt(0); // reserved1
            bos.WriteInt(0); // reserved2
            String cspName = CspName;
            if (cspName == null) cspName = CipherProvider.cipherProviderName;
            bos.Write(StringUtil.GetToUnicodeLE(cspName));
            bos.WriteShort(0);
            int headerSize = bos.WriteIndex - startIdx - LittleEndianConsts.INT_SIZE;
            sizeOutput.WriteInt(headerSize);
        }
    }

}
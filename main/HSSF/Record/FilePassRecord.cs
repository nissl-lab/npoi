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

namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.Util;
    using System.Text;
    using System.Diagnostics;
    using System.Collections.Generic;
    using System.IO;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.BinaryRC4;
    using NPOI.POIFS.Crypt.CryptoAPI;
    using NPOI.POIFS.Crypt.XOR;

    /**
     * Title: File Pass Record (0x002F) <p>
     *
     * Description: Indicates that the record After this record are encrypted.
     */
    public class FilePassRecord : StandardRecord
    {
        public const short sid = 0x002F;
        private const int ENCRYPTION_XOR = 0;
        private const int ENCRYPTION_OTHER = 1;

        private readonly int encryptionType;
        private readonly EncryptionInfo encryptionInfo;

        public FilePassRecord(EncryptionMode encryptionMode)
        {
            encryptionType = (encryptionMode == EncryptionMode.XOR) ? ENCRYPTION_XOR : ENCRYPTION_OTHER;
            encryptionInfo = new EncryptionInfo(encryptionMode);
        }

        public FilePassRecord(RecordInputStream input)
        {
            encryptionType = input.ReadUShort();

            EncryptionMode preferredMode;
            switch (encryptionType)
            {
                case ENCRYPTION_XOR:
                    preferredMode = EncryptionMode.XOR;
                    break;
                case ENCRYPTION_OTHER:
                    preferredMode = EncryptionMode.CryptoAPI;
                    break;
                default:
                    throw new EncryptedDocumentException("invalid encryption type");
            }

            try
            {
                encryptionInfo = new EncryptionInfo(input, preferredMode);
            }
            catch (IOException e)
            {
                throw new EncryptedDocumentException(e);
            }
        }

        // Replace the switch statement in Serialize with if/else, since switch requires constant values for case labels
        public override void Serialize(ILittleEndianOutput output)
        {
            output.WriteShort(encryptionType);

            byte[] data = new byte[1024];
            var bos = new LittleEndianByteArrayOutputStream(data, 0);

            if (encryptionInfo.EncryptionMode == EncryptionMode.XOR)
            {
                ((XOREncryptionHeader)encryptionInfo.Header).Write(bos);
                ((XOREncryptionVerifier)encryptionInfo.Verifier).Write(bos);
            }
            else if (encryptionInfo.EncryptionMode == EncryptionMode.BinaryRC4)
            {
                output.WriteShort(encryptionInfo.VersionMajor);
                output.WriteShort(encryptionInfo.VersionMinor);
                ((BinaryRC4EncryptionHeader)encryptionInfo.Header).Write(bos);
                ((BinaryRC4EncryptionVerifier)encryptionInfo.Verifier).Write(bos);
            }
            else if (encryptionInfo.EncryptionMode == EncryptionMode.CryptoAPI)
            {
                output.WriteShort(encryptionInfo.VersionMajor);
                output.WriteShort(encryptionInfo.VersionMinor);
                output.WriteInt(encryptionInfo.EncryptionFlags);
                ((CryptoAPIEncryptionHeader)encryptionInfo.Header).Write(bos);
                ((CryptoAPIEncryptionVerifier)encryptionInfo.Verifier).Write(bos);
            }
            else
            {
                throw new EncryptedDocumentException("not supported");
            }

            output.Write(data, 0, bos.WriteIndex);
        }

        protected override int DataSize
        {
            get
            {
                using (MemoryStream bos = new MemoryStream())
                using (var leos = new LittleEndianOutputStream(bos))
                {
                    Serialize(leos);
                    return (int)bos.Length;
                }
            }
        }

        public EncryptionInfo GetEncryptionInfo()
        {
            return encryptionInfo;
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override String ToString()
        {
            return $"[FILEPASS] encryptionType={encryptionType} encryptionInfo={encryptionInfo}";
        }
    }

}
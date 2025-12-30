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

namespace NPOI.POIFS.Crypt.XOR
{
    using System;
    using System.IO;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public class XOREncryptor : Encryptor
    {
        private IEncryptionInfoBuilder builder;
        private ISecretKey secretKey;

        public XOREncryptor(IEncryptionInfoBuilder builder)
        {
            this.builder = builder;
        }

        protected XOREncryptor(XOREncryptor other) : base(other)
        {
            builder = other.builder;
            // secretKey is immutable
            secretKey = other.secretKey;
        }

        public override void ConfirmPassword(String password)
        {
            int keyComp = CryptoFunctions.CreateXorKey1(password);
            int verifierComp = CryptoFunctions.CreateXorVerifier1(password);
            byte[] xorArray = CryptoFunctions.CreateXorArray1(password);

            byte[] shortBuf = new byte[2];
            XOREncryptionVerifier ver = (XOREncryptionVerifier)builder.GetVerifier();
            LittleEndian.PutUShort(shortBuf, 0, keyComp);
            ver.SetEncryptedKey(shortBuf);
            LittleEndian.PutUShort(shortBuf, 0, verifierComp);
            ver.SetEncryptedVerifier(shortBuf);
            SetSecretKey(new SecretKeySpec(xorArray, "XOR"));
        }

        public override void ConfirmPassword(String password, byte[] keySpec,
                            byte[] keySalt, byte[] verifier, byte[] verifierSalt,
                            byte[] integritySalt)
        {
            ConfirmPassword(password);
        }

        public override OutputStream GetDataStream(DirectoryNode dir)
        {
            return new XORCipherOutputStream(this);
        }

        protected int GetKeySizeInBytes()
        {
            return -1;
        }

        public XOREncryptor Copy()
        {
            return new XOREncryptor(this);
        }

        public void SetSecretKey(ISecretKey secretKey)
        {
            this.secretKey = secretKey;
        }

        public ISecretKey GetSecretKey()
        {
            return secretKey;
        }

        // Change the accessibility of XORCipherOutputStream from private to public
        public class XORCipherOutputStream : ByteArrayOutputStream, IDisposable
        {
            private XOREncryptor encryptor;
            private byte[] oneByte = { 0 };
            private int recordStart;
            private int recordEnd;
            private byte[] xorArray;

            public XORCipherOutputStream(XOREncryptor encryptor) : base()
            {
                this.encryptor = encryptor;
                this.xorArray = encryptor.GetSecretKey().GetEncoded();
            }

            public byte[] GetBuf()
            {
                return base.ToByteArray();
            }

            public void SetSize(long count)
            {
                base.SetLength(count);
            }

            public void SetNextRecordSize(int recordSize, bool isPlain)
            {
                if (recordEnd > 0 && !isPlain)
                {
                    // Process last record if needed
                }
                recordStart = (int)Length + 4;
                recordEnd = recordStart + recordSize;
            }

            public override void Write(int b)
            {
                oneByte[0] = (byte)b;
                EncryptByte(oneByte, 0, 1);
                base.Write(oneByte);
            }

            public override void Write(byte[] b, int off, int len)
            {
                // Create a copy to avoid modifying the original array
                byte[] encrypted = new byte[len];
                Array.Copy(b, off, encrypted, 0, len);
                EncryptBytes(encrypted, 0, len);
                base.Write(encrypted, 0, len);
            }

            private void EncryptByte(byte[] data, int offset, int length)
            {
                EncryptBytes(data, offset, length);
            }

            private void EncryptBytes(byte[] data, int offset, int length)
            {
                if (xorArray == null) return;

                /*
                 * From: http://social.msdn.microsoft.com/Forums/en-US/3dadbed3-0e68-4f11-8b43-3a2328d9ebd5
                 *
                 * The initial value for XorArrayIndex is as follows:
                 * XorArrayIndex = (FileOffset + Data.Length) % 16
                 *
                 * The FileOffset variable in this context is the stream offset into the Workbook stream at
                 * the time we are about to write each of the bytes of the record data.
                 * This (the value) is then incremented after each byte is written.
                 */
                int xorArrayIndex = (int)(Length + offset) % 16;

                for (int i = 0; i < length; i++)
                {
                    byte value = data[offset + i];
                    value ^= xorArray[xorArrayIndex];
                    value = RotateLeft(value, 8 - 3);
                    data[offset + i] = value;
                    xorArrayIndex = (xorArrayIndex + 1) & 0x0F;
                }
            }

            private static byte RotateLeft(byte bits, int shift)
            {
                return (byte)(((bits & 0xff) << shift) | ((bits & 0xff) >> (8 - shift)));
            }
        }
    }
}
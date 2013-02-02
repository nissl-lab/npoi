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
using System.Security.Cryptography;

using NPOI.POIFS.FileSystem;
using NPOI.Util;

namespace NPOI.POIFS.Crypt
{
    public class EcmaDecryptor : Decryptor
    {
        private EncryptionInfo info;
        private byte[] passwordHash;
        private long _length = -1;
        public EcmaDecryptor(EncryptionInfo info)
        {
            this.info = info;
        }

        private byte[] GenerateKey(int block)
        {
            try
            {
                HashAlgorithm sha1 = HashAlgorithm.Create("SHA1");

                //sha1.update(passwordHash);

                byte[] blockValue = new byte[4];
                LittleEndian.PutInt(blockValue, 0, block);
                byte[] temp = new byte[blockValue.Length + passwordHash.Length];
                Array.Copy(passwordHash, temp, passwordHash.Length);
                Array.Copy(blockValue, 0, temp, passwordHash.Length, blockValue.Length);
                byte[] finalHash = sha1.ComputeHash(temp); //sha1.digest(blockValue);

                int reqiredKeyLength = (info.Header.KeySize) / 8;

                byte[] buff = new byte[64];

                for (int i = 0; i < buff.Length; i++)
                    buff[i] = (byte)0x36;

                for (int i = 0; i < finalHash.Length; i++)
                    buff[i] = (byte)(buff[i] ^ finalHash[i]);

                //sha1.Clear();

                byte[] x1 = sha1.ComputeHash(buff);
                for (int i = 0; i < buff.Length; i++)
                    buff[i] = (byte)0x5c;
                for (int i = 0; i < finalHash.Length; i++)
                    buff[i] = (byte)(buff[i] ^ finalHash[i]);

                //sha1.Clear();

                byte[] x2 = sha1.ComputeHash(buff);
                byte[] x3 = new byte[x1.Length + x2.Length];

                System.Array.Copy(x1, 0, x3, 0, x1.Length);

                System.Array.Copy(x2, 0, x3, x1.Length, x2.Length);

                return TruncateOrPad(x3, reqiredKeyLength);

            }
            catch (System.Security.Cryptography.CryptographicException ex)
            {
                throw ex;
            }
        }

        public override bool VerifyPassword(string password)
        {
            passwordHash = HashPassword(info, password);

            SymmetricAlgorithm cipher = GetCipher();
            byte[] verifier; // byte[] verifier = cipher.doFinal(info.getVerifier().getVerifier());
            //using (MemoryStream fStream = new MemoryStream(info.Verifier.Verifier))
            //{
            //    using (CryptoStream cStream = new CryptoStream(fStream, cipher.CreateDecryptor(cipher.Key, cipher.IV),
            //        CryptoStreamMode.Read))
            //    {
            //        verifier = new byte[cStream.Length];//??
            //        cStream.Read(verifier, 0, verifier.Length);
            //    }
            //}
            verifier = Decrypt(cipher, info.Verifier.Verifier);
            HashAlgorithm sha1 = HashAlgorithm.Create("SHA1");
            byte[] calcVerifierHash = sha1.ComputeHash(verifier);
            byte[] temp;
            //using (MemoryStream fStream = new MemoryStream(info.Verifier.VerifierHash))
            //{
            //    using (CryptoStream cStream = new CryptoStream(fStream, cipher.CreateDecryptor(cipher.Key, cipher.IV),
            //        CryptoStreamMode.Read))
            //    {
            //        temp = new byte[cStream.Length];
            //        cStream.Read(temp, 0, verifier.Length);
            //    }
            //}
            temp = Decrypt(cipher, info.Verifier.VerifierHash);
            byte[] verifierHash = TruncateOrPad(calcVerifierHash, calcVerifierHash.Length);

            return Arrays.Equals(calcVerifierHash, verifierHash);

        }

        private byte[] TruncateOrPad(byte[] source, int length)
        {
            byte[] result = new byte[length];
            System.Array.Copy(source, 0, result, 0, Math.Min(length, source.Length));

            if (length > source.Length)
            {
                for (int i = source.Length; i < length; i++)
                    result[i] = 0;
            }

            return result;
        }
        

        private SymmetricAlgorithm GetCipher()
        {
            byte[] key = GenerateKey(0);
            //System.Security.Cryptography

            SymmetricAlgorithm cipher = SymmetricAlgorithm.Create();  //  AES/ECB/NoPadding
            cipher.Mode = CipherMode.ECB;
            cipher.Padding = PaddingMode.None;
            cipher.Key = key;
            
            //cipher.init(Cipher.DECRYPT_MODE, skey); Leon
            return cipher;
        }

        public override Stream GetDataStream(DirectoryNode dir)
        {
            DocumentInputStream dis = dir.CreateDocumentInputStream("EncryptedPackage");
            _length = dis.ReadLong();
            SymmetricAlgorithm cipher=GetCipher();
            return new CryptoStream(dis, cipher.CreateDecryptor(cipher.Key, cipher.IV), CryptoStreamMode.Read);
        }

        public override long Length
        {
            get
            {
                if (_length == -1) throw new InvalidOperationException("EcmaDecryptor.getDataStream() was not called");
                return _length;
            }
        }
    }
}

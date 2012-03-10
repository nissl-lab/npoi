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
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security;
using System.Security.Cryptography;

using NPOI.POIFS.FileSystem;
using NPOI.Util;

namespace NPOI.POIFS.Crypt
{
    public class EcmaDecryptor : Decryptor
    {
        private EncryptionInfo info;
        private byte[] passwordHash;

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
                LittleEndian.PutInt(blockValue, block);
                byte[] finalHash = sha1.ComputeHash(blockValue); //sha1.digest(blockValue);

                int reqiredKeyLength = (info.GetHeader().KeySize) / 8;

                byte[] buff = new byte[64];

            //    for (int i = 0; i < buff.Length; i++)
                  //  buff[i] = (byte)0x36;

                for (int i = 0; i < finalHash.Length; i++)
                    buff[i] = (byte)(0x36^ finalHash[i]);

                sha1.Clear();

                byte[] x1 = sha1.ComputeHash(buff);

                for (int i = 0; i < buff.Length; i++)
                    buff[i] = (byte)(0x5c ^ (finalHash[i]));

                sha1.Clear();

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

            byte[] verifier = cipher.Key; // byte[] verifier = cipher.doFinal(info.getVerifier().getVerifier());

            HashAlgorithm sha1 = HashAlgorithm.Create("SHA1");
            byte[] calcVerifierHash = sha1.ComputeHash(verifier);

            byte[] verifierHash = TruncateOrPad(cipher.Key, calcVerifierHash.Length);

            return Array.Equals(calcVerifierHash, verifierHash);

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
            SymmetricAlgorithm cipher = SymmetricAlgorithm.Create("AES/ECB/NoPadding");
            cipher.Mode = CipherMode.ECB;
            cipher.Key = key;
            //cipher.init(Cipher.DECRYPT_MODE, skey); Leon
            return cipher;
        }

        public override Stream GetDataStream(DirectoryNode dir)
        {
            DocumentReader dr = dir.CreateDocumentInputStream("EncryptedPackage");
            long size = dr.ReadLong();

            return new CryptoStream(dr, /*GetCipher()*/ null, CryptoStreamMode.Write);
        }
    }
}

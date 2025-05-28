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

namespace NPOI.POIFS.Crypt.BinaryRC4
{
    using System;
    using System.IO;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.CryptoAPI;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public class BinaryRC4Decryptor : Decryptor
    {
        private long _length = -1L;

        private sealed class BinaryRC4CipherInputStream : ChunkedCipherInputStream
        {
            protected override Cipher InitCipherForBlock(Cipher existing, int block)
            {
                return BinaryRC4Decryptor.InitCipherForBlock(existing, block,
                    decryptor.builder, decryptor.GetSecretKey(), Cipher.DECRYPT_MODE);
            }

            public BinaryRC4CipherInputStream(DocumentInputStream stream, long size, Decryptor decryptor)
                : base(stream, size, 512, decryptor.builder, decryptor)
            {
                this.decryptor = decryptor;
            }
        }

        protected internal BinaryRC4Decryptor(BinaryRC4EncryptionInfoBuilder builder)
            : base(builder)
        {

        }

        public override bool VerifyPassword(String password)
        {
            EncryptionVerifier ver = builder.GetVerifier();
            ISecretKey skey = GenerateSecretKey(password, ver);
            try
            {
                Cipher cipher = InitCipherForBlock(null, 0, builder, skey, Cipher.DECRYPT_MODE);
                byte[] encryptedVerifier = ver.EncryptedVerifier;
                byte[] verifier = new byte[encryptedVerifier.Length];
                cipher.Update(encryptedVerifier, 0, encryptedVerifier.Length, verifier);
                SetVerifier(verifier);
                byte[] encryptedVerifierHash = ver.EncryptedVerifierHash;
                byte[] verifierHash = cipher.DoFinal(encryptedVerifierHash);
                HashAlgorithm hashAlgo = ver.HashAlgorithm;
                MessageDigest hashAlg = CryptoFunctions.GetMessageDigest(hashAlgo);
                byte[] calcVerifierHash = hashAlg.Digest(verifier);
                if (Arrays.Equals(calcVerifierHash, verifierHash))
                {
                    SetSecretKey(skey);
                    return true;
                }
            }
            catch (Exception e)
            {
                throw new EncryptedDocumentException(e);
            }
            return false;
        }

        protected internal static Cipher InitCipherForBlock(Cipher cipher, int block,
            IEncryptionInfoBuilder builder, ISecretKey skey, int encryptMode)
        {
            EncryptionVerifier ver = builder.GetVerifier();
            HashAlgorithm hashAlgo = ver.HashAlgorithm;
            byte[] blockKey = new byte[4];
            LittleEndian.PutUInt(blockKey, 0, block);
            byte[] encKey = CryptoFunctions.GenerateKey(skey.GetEncoded(), hashAlgo, blockKey, 16);
            ISecretKey key = new SecretKeySpec(encKey, skey.GetAlgorithm());
            if (cipher == null)
            {
                EncryptionHeader em = builder.GetHeader();
                cipher = CryptoFunctions.GetCipher(key, em.CipherAlgorithm, null, null, encryptMode);
            }
            else
            {
                cipher.Init(encryptMode, key);
            }
            return cipher;
        }

        protected internal static ISecretKey GenerateSecretKey(String password,
                EncryptionVerifier ver)
        {
            if (password.Length > 255)
                password = password.Substring(0, 255);
            HashAlgorithm hashAlgo = ver.HashAlgorithm;
            MessageDigest hashAlg = CryptoFunctions.GetMessageDigest(hashAlgo);
            byte[] hash = hashAlg.Digest(StringUtil.GetToUnicodeLE(password));
            byte[] salt = ver.Salt;
            hashAlg.Reset();
            for (int i = 0; i < 16; i++)
            {
                hashAlg.Update(hash, 0, 5);
                hashAlg.Update(salt);
            }

            hash = new byte[5];
            Array.Copy(hashAlg.Digest(), 0, hash, 0, 5);
            ISecretKey skey = new SecretKeySpec(hash, ver.CipherAlgorithm.jceId);
            return skey;
        }

        public override InputStream GetDataStream(DirectoryNode dir)
        {
            DocumentInputStream dis = dir.CreateDocumentInputStream(DEFAULT_POIFS_ENTRY);
            _length = dis.ReadLong();
            BinaryRC4CipherInputStream cipherStream = new BinaryRC4CipherInputStream(dis, _length, this);
            //return cipherStream.GetStream();
            throw new NotImplementedException("BinaryRC4CipherInputStream should be derived from InputStream");
        }

        public override long GetLength()
        {
            if (_length == -1L)
            {
                throw new InvalidOperationException("Decryptor.DataStream was not called");
            }

            return _length;
        }
    }

}
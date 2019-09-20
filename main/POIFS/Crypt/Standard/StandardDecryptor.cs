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
    using System.Diagnostics;
    using System.IO;
    using NPOI.POIFS.Crypt;

    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    /**
     */
    public class StandardDecryptor : Decryptor {
        private long _length = -1;

        internal StandardDecryptor(IEncryptionInfoBuilder builder)
            : base(builder)
        {
            
        }

        public override bool VerifyPassword(String password) {
            EncryptionVerifier ver = builder.GetVerifier();
            ISecretKey skey = GenerateSecretKey(password, ver, GetKeySizeInBytes());
            Cipher cipher = GetCipher(skey);

            try {
                byte[] encryptedVerifier = ver.EncryptedVerifier;
                byte[] verifier = cipher.DoFinal(encryptedVerifier);
                SetVerifier(verifier);
                MessageDigest sha1 = CryptoFunctions.GetMessageDigest(ver.HashAlgorithm);
                byte[] calcVerifierHash = sha1.Digest(verifier);
                byte[] encryptedVerifierHash = ver.EncryptedVerifierHash;
                byte[] decryptedVerifierHash = cipher.DoFinal(encryptedVerifierHash);

                // see 2.3.4.9 Password Verification (Standard Encryption)
                // ... The number of bytes used by the encrypted Verifier hash MUST be 32 ...
                // TODO: check and Trim/pad the hashes to 32
                byte[] verifierHash = Arrays.CopyOf(decryptedVerifierHash, calcVerifierHash.Length);

                if (Arrays.Equals(calcVerifierHash, verifierHash)) {
                    SetSecretKey(skey);
                    return true;
                } else {
                    return false;
                }
            } catch (Exception e) {
                throw new EncryptedDocumentException(e);
            }
        }

        protected internal static ISecretKey GenerateSecretKey(String password, EncryptionVerifier ver, int keySize) {
            HashAlgorithm hashAlgo = ver.HashAlgorithm;

            byte[] pwHash = CryptoFunctions.HashPassword(password, hashAlgo, ver.Salt, ver.SpinCount);

            byte[] blockKey = new byte[4];
            LittleEndian.PutInt(blockKey, 0, 0);

            byte[] finalHash = CryptoFunctions.GenerateKey(pwHash, hashAlgo, blockKey, hashAlgo.hashSize);
            byte[] x1 = FillAndXor(finalHash, (byte)0x36);
            byte[] x2 = FillAndXor(finalHash, (byte)0x5c);

            byte[] x3 = new byte[x1.Length + x2.Length];
            Array.Copy(x1, 0, x3, 0, x1.Length);
            Array.Copy(x2, 0, x3, x1.Length, x2.Length);

            byte[] key = Arrays.CopyOf(x3, keySize);

            throw new NotImplementedException();
            //SecretKey skey = new SecretKeySpec(key, ver.CipherAlgorithm.jceId);
            //return skey;
        }

        protected static byte[] FillAndXor(byte[] hash, byte FillByte) {
            byte[] buff = new byte[64];
            Arrays.Fill(buff, FillByte);

            for (int i = 0; i < hash.Length; i++) {
                buff[i] = (byte)(buff[i] ^ hash[i]);
            }

            throw new NotImplementedException();
            //MessageDigest sha1 = CryptoFunctions.GetMessageDigest(HashAlgorithm.sha1);
            //return sha1.Digest(buff);
        }

        private Cipher GetCipher(ISecretKey key) {
            EncryptionHeader em = builder.GetHeader();
            ChainingMode cm = em.ChainingMode;
            Debug.Assert(cm == ChainingMode.ecb);
            throw new NotImplementedException();
            //return CryptoFunctions.GetCipher(key, em.GetCipherAlgorithm(), cm, null, Cipher.DECRYPT_MODE);
        }

        public override Stream GetDataStream(DirectoryNode dir) {
            DocumentInputStream dis = dir.CreateDocumentInputStream(Encryptor.DEFAULT_POIFS_ENTRY);

            _length = dis.ReadLong();

            // limit wrong calculated ole entries - (bug #57080)
            // standard encryption always uses aes encoding, so blockSize is always 16 
            // http://stackoverflow.com/questions/3283787/size-of-data-after-aes-encryption
            int blockSize = builder.GetHeader().CipherAlgorithm.blockSize;
            long cipherLen = (_length / blockSize + 1) * blockSize;
            Cipher cipher = GetCipher(GetSecretKey());

            throw new NotImplementedException();
            //InputStream boundedDis = new BoundedInputStream(dis, cipherLen);
            //return new BoundedInputStream(new CipherInputStream(boundedDis, cipher), _length);
        }

        /**
         * @return the length of the stream returned by {@link #getDataStream(DirectoryNode)}
         */
        public override long GetLength() {
            if (_length == -1) throw new InvalidOperationException("Decryptor.DataStream was not called");
            return _length;
        }
    }

}
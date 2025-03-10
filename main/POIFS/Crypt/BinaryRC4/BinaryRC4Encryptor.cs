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

    using NPOI.POIFS.Crypt.Standard;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public class BinaryRC4Encryptor : Encryptor
    {

        private BinaryRC4EncryptionInfoBuilder builder;

        protected class BinaryRC4CipherOutputStream : ChunkedCipherOutputStream
        {
            protected override Cipher InitCipherForBlock(Cipher cipher, int block, bool lastChunk)
            {
                return BinaryRC4Decryptor.InitCipherForBlock(cipher, block, builder, encryptor.GetSecretKey(), Cipher.ENCRYPT_MODE);
            }

            protected override void CalculateChecksum(FileInfo file, int i)
            {
            }

            protected override void CreateEncryptionInfoEntry(DirectoryNode dir, FileInfo tmpFile)
            {
                ((BinaryRC4Encryptor)encryptor).CreateEncryptionInfoEntry(dir);
            }

            public BinaryRC4CipherOutputStream(DirectoryNode dir, BinaryRC4EncryptionInfoBuilder builder, BinaryRC4Encryptor encryptor)
                : base(dir, 512, builder, encryptor)
            {
                
            }
        }

        protected internal BinaryRC4Encryptor(BinaryRC4EncryptionInfoBuilder builder)
        {
            this.builder = builder;
        }

        public override void ConfirmPassword(String password)
        {
            Random r = new Random();
            byte[] salt = new byte[16];
            byte[] verifier = new byte[16];
            r.NextBytes(salt);
            r.NextBytes(verifier);
            ConfirmPassword(password, null, null, verifier, salt, null);
        }

        public override void ConfirmPassword(String password, byte[] keySpec,
                byte[] keySalt, byte[] verifier, byte[] verifierSalt,
                byte[] integritySalt)
        {
            BinaryRC4EncryptionVerifier ver = builder.GetVerifier();
            ver.SetSalt(verifierSalt);
            ISecretKey skey = BinaryRC4Decryptor.GenerateSecretKey(password, ver);
            SetSecretKey(skey);
            try
            {
                Cipher cipher = BinaryRC4Decryptor.InitCipherForBlock(null, 0, builder, skey, Cipher.ENCRYPT_MODE);
                byte[] encryptedVerifier = new byte[16];
                cipher.Update(verifier, 0, 16, encryptedVerifier);
                ver.EncryptedVerifier = (encryptedVerifier);
                HashAlgorithm hashAlgo = ver.HashAlgorithm;
                MessageDigest hashAlg = CryptoFunctions.GetMessageDigest(hashAlgo);
                byte[] calcVerifierHash = hashAlg.Digest(verifier);
                byte[] encryptedVerifierHash = cipher.DoFinal(calcVerifierHash);
                ver.EncryptedVerifierHash = (encryptedVerifierHash);
            }
            catch (Exception e)
            {
                throw new EncryptedDocumentException("Password Confirmation failed", e);
            }
        }

        public override OutputStream GetDataStream(DirectoryNode dir)
        {
            Stream countStream = new BinaryRC4CipherOutputStream(dir, builder, this).out1;
            throw new NotImplementedException("BinaryRC4CipherOutputStream should be derived from OutputStream");
            //return countStream;
        }

        protected int GetKeySizeInBytes()
        {
            return builder.GetHeader().KeySize / 8;
        }

        protected internal void CreateEncryptionInfoEntry(DirectoryNode dir)
        {
            DataSpaceMapUtils.AddDefaultDataSpace(dir);
            EncryptionInfo info = builder.GetEncryptionInfo();
            BinaryRC4EncryptionHeader header = builder.GetHeader();
            BinaryRC4EncryptionVerifier verifier = builder.GetVerifier();
            EncryptionRecord er = new EncryptionRecordInternal(info, header, verifier);
            
            DataSpaceMapUtils.CreateEncryptionEntry(dir, "EncryptionInfo", er);
        }

        private sealed class EncryptionRecordInternal : EncryptionRecord
        {
            public EncryptionRecordInternal(EncryptionInfo info,
                BinaryRC4EncryptionHeader header, BinaryRC4EncryptionVerifier verifier)
            {
                this.info = info;
                this.header = header;
                this.verifier = verifier;
            }
            EncryptionInfo info;
            BinaryRC4EncryptionHeader header;
            BinaryRC4EncryptionVerifier verifier;
            public void Write(LittleEndianByteArrayOutputStream bos)
            {
                bos.WriteShort(info.VersionMajor);
                bos.WriteShort(info.VersionMinor);
                header.Write(bos);
                verifier.Write(bos);
            }
        }
    }

}
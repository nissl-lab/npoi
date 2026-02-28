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
        private int _chunkSize = 512;

        public BinaryRC4Encryptor(BinaryRC4EncryptionInfoBuilder builder)
        {
            this.builder = builder;
        }

        protected BinaryRC4Encryptor(BinaryRC4Encryptor other) : base(other)
        {
            this.builder = other.builder;
            _chunkSize = other._chunkSize;
        }

        public override OutputStream GetDataStream(DirectoryNode dir)
        {
            return new BinaryRC4CipherOutputStream(dir, builder, this);
        }

        public override ChunkedCipherOutputStream GetDataStream(OutputStream stream, int initialOffset)
        {
            return new BinaryRC4CipherOutputStream(stream, builder, this);
        }

        protected class BinaryRC4CipherOutputStream : ChunkedCipherOutputStream
        {
            public BinaryRC4CipherOutputStream(OutputStream stream, IEncryptionInfoBuilder builder, BinaryRC4Encryptor binaryRC4Encryptor)
                : base(stream, binaryRC4Encryptor._chunkSize, builder, binaryRC4Encryptor)
            {
            }

            //If not, wire your POIFS packaging where you create the output stream.
            public BinaryRC4CipherOutputStream(DirectoryNode dir, IEncryptionInfoBuilder builder, BinaryRC4Encryptor binaryRC4Encryptor)
                : base(dir, binaryRC4Encryptor._chunkSize, builder, binaryRC4Encryptor)
            {
            }

            protected override Cipher InitCipherForBlock(Cipher cipher, int block, bool lastChunk)
            {
                // Reinitialize for given block in ENCRYPT mode.
                return BinaryRC4Decryptor.InitCipherForBlock(cipher, block,
                    builder,
                    encryptor.GetSecretKey(),
                    Cipher.ENCRYPT_MODE);
            }

            protected override void CalculateChecksum(FileInfo fileOut, int oleStreamSize)
            {
            }

            protected override void CreateEncryptionInfoEntry(DirectoryNode dir, FileInfo tmpFile)
            {
                ((BinaryRC4Encryptor)encryptor).CreateEncryptionInfoEntry(dir);
            }
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
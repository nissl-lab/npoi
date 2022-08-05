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

namespace NPOI.POIFS.Crypt.CryptoAPI
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public class CryptoAPIDecryptor : Decryptor {

        private long _length;

        private class SeekableMemoryStream : MemoryStream {
            Cipher cipher;
            byte[] oneByte = { 0 };

            public void Seek(int pos) {
                if (pos > Length) {
                    throw new IndexOutOfRangeException(string.Format( "seek position({0}) is greater than stream length({1})", pos, Length));
                }

                //this.pos = pos;
                //mark = pos;
                throw new NotImplementedException();
            }

            public void SetBlock(int block) {
                //cipher = InitCipherForBlock(cipher, block);
                throw new NotImplementedException();
            }

            public int Read() {
                int ch = base.ReadByte();
                if (ch == -1) return -1;
                oneByte[0] = (byte)ch;
                try {
                    cipher.Update(oneByte, 0, 1, oneByte);
                } catch (Exception e) {
                    throw new EncryptedDocumentException(e);
                }
                return oneByte[0];
            }

            public override int Read(byte[] b, int off, int len) {
                int readLen = base.Read(b, off, len);
                if (readLen == -1) return 0;
                try {
                    cipher.Update(b, off, readLen, b, off);
                } catch (Exception e) {
                    throw new EncryptedDocumentException(e);
                }
                return readLen;
            }

            public SeekableMemoryStream(byte[] buf)
                :base(buf)
            {
                //cipher = InitCipherForBlock(null, 0);
                throw new NotImplementedException();
            }
        }

        internal class StreamDescriptorEntry {
            internal static BitField flagStream = BitFieldFactory.GetInstance(1);

            internal int streamOffset;
            internal int streamSize;
            internal int block;
            internal int flags;
            internal int reserved2;
            internal String streamName;
        }

        protected internal CryptoAPIDecryptor(CryptoAPIEncryptionInfoBuilder builder) : base(builder)
        {
            
            _length = -1L;
        }

        public override bool VerifyPassword(String password) {
            EncryptionVerifier ver = builder.GetVerifier();
            ISecretKey skey = GenerateSecretKey(password, ver);
            try {
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
                if (Arrays.Equals(calcVerifierHash, verifierHash)) {
                    SetSecretKey(skey);
                    return true;
                }
            } catch (Exception e) {
                throw new EncryptedDocumentException(e);
            }
            return false;
        }

        /**
         * Initializes a cipher object for a given block index for decryption
         *
         * @param cipher may be null, otherwise the given instance is reset to the new block index
         * @param block the block index, e.g. the persist/slide id (hslf)
         * @return a new cipher object, if cipher was null, otherwise the reInitialized cipher
         * @throws GeneralSecurityException
         */
        public Cipher InitCipherForBlock(Cipher cipher, int block)
        {
            return InitCipherForBlock(cipher, block, builder, GetSecretKey(), Cipher.DECRYPT_MODE);
        }

        protected internal static Cipher InitCipherForBlock(Cipher cipher, int block,
            IEncryptionInfoBuilder builder, ISecretKey skey, int encryptMode)
        {
            EncryptionVerifier ver = builder.GetVerifier();
            HashAlgorithm hashAlgo = ver.HashAlgorithm;
            byte[] blockKey = new byte[4];
            LittleEndian.PutUInt(blockKey, 0, block);
            MessageDigest hashAlg = CryptoFunctions.GetMessageDigest(hashAlgo);
            hashAlg.Update(skey.GetEncoded());
            byte[] encKey = hashAlg.Digest(blockKey);
            EncryptionHeader header = builder.GetHeader();
            int keyBits = header.KeySize;
            encKey = CryptoFunctions.GetBlock0(encKey, keyBits / 8);
            if (keyBits == 40) {
                encKey = CryptoFunctions.GetBlock0(encKey, 16);
            }
            ISecretKey key = new SecretKeySpec(encKey, skey.GetAlgorithm());
            if (cipher == null) {
                cipher = CryptoFunctions.GetCipher(key, header.CipherAlgorithm, null, null, encryptMode);
            } else {
                cipher.Init(encryptMode, key);
            }
            return cipher;
        }

        protected internal static ISecretKey GenerateSecretKey(String password, EncryptionVerifier ver) {
            if (password.Length > 255) {
                password = password.Substring(0, 255);
            }
            HashAlgorithm hashAlgo = ver.HashAlgorithm;
            MessageDigest hashAlg = CryptoFunctions.GetMessageDigest(hashAlgo);
            hashAlg.Update(ver.Salt);
            byte[] hash = hashAlg.Digest(StringUtil.GetToUnicodeLE(password));
            ISecretKey skey = new SecretKeySpec(hash, ver.CipherAlgorithm.jceId);
            return skey;
        }

        /**
         * Decrypt the Document-/SummaryInformation and other optionally streams.
         * Opposed to other crypto modes, cryptoapi is record based and can't be used
         * to stream-decrypt a whole file
         * 
         * @see <a href="http://msdn.microsoft.com/en-us/library/dd943321(v=office.12).aspx">2.3.5.4 RC4 CryptoAPI Encrypted Summary Stream</a>
         */

        public override InputStream GetDataStream(DirectoryNode dir)
        {
            NPOIFSFileSystem fsOut = new NPOIFSFileSystem();
            DocumentNode es = (DocumentNode)dir.GetEntry("EncryptedSummary");
            DocumentInputStream dis = dir.CreateDocumentInputStream(es);
            MemoryStream bos = new MemoryStream();
            IOUtils.Copy(dis, bos);
            dis.Close();
            SeekableMemoryStream sbis = new SeekableMemoryStream(bos.ToArray());
            LittleEndianInputStream leis = new LittleEndianInputStream(sbis);
            int streamDescriptorArrayOffset = (int)leis.ReadUInt();
            int streamDescriptorArraySize = (int)leis.ReadUInt();
            sbis.Seek(streamDescriptorArrayOffset - 8, SeekOrigin.Current);// sbis.Skip(streamDescriptorArrayOffset - 8);
            
            sbis.SetBlock(0);
            int encryptedStreamDescriptorCount = (int)leis.ReadUInt();
            StreamDescriptorEntry[] entries = new StreamDescriptorEntry[encryptedStreamDescriptorCount];
            for (int i = 0; i < encryptedStreamDescriptorCount; i++) {
                StreamDescriptorEntry entry = new StreamDescriptorEntry();
                entries[i] = entry;
                entry.streamOffset = (int)leis.ReadUInt();
                entry.streamSize = (int)leis.ReadUInt();
                entry.block = leis.ReadUShort();
                int nameSize = leis.ReadUByte();
                entry.flags = leis.ReadUByte();
                bool IsStream = StreamDescriptorEntry.flagStream.IsSet(entry.flags);
                entry.reserved2 = leis.ReadInt();
                entry.streamName = StringUtil.ReadUnicodeLE(leis, nameSize);
                leis.ReadShort();
                Debug.Assert(entry.streamName.Length == nameSize);
            }

            foreach (StreamDescriptorEntry entry in entries) {
                sbis.Seek(entry.streamOffset);
                sbis.SetBlock(entry.block);
                Stream is1 = new BufferedStream(sbis, entry.streamSize);
                fsOut.CreateDocument(is1, entry.streamName);
            }

            leis.Close();
            sbis = null;
            
            bos.Seek(0, SeekOrigin.Begin); //bos.Reset();
            fsOut.WriteFileSystem(bos);
            fsOut.Close();
            _length = bos.Length;
            ByteArrayInputStream bis = new ByteArrayInputStream(bos.ToArray());
            throw new NotImplementedException("ByteArrayInputStream should be derived from InputStream");
        }

        /**
         * @return the length of the stream returned by {@link #getDataStream(DirectoryNode)}
         */
        public override long GetLength() {
            if (_length == -1L) {
                throw new InvalidOperationException("Decryptor.DataStream was not called");
            }
            return _length;
        }
    }


}

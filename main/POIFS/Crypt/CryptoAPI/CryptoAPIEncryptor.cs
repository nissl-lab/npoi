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
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using NPOI.HPSF;

    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.Crypt.Standard;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public class CryptoAPIEncryptor : Encryptor
    {
        private CryptoAPIEncryptionInfoBuilder builder;

        protected internal CryptoAPIEncryptor(CryptoAPIEncryptionInfoBuilder builder)
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
            Debug.Assert(verifier != null && verifierSalt != null);
            CryptoAPIEncryptionVerifier ver = builder.GetVerifier();
            ver.SetSalt(verifierSalt);
            ISecretKey skey = CryptoAPIDecryptor.GenerateSecretKey(password, ver);
            SetSecretKey(skey);
            try
            {
                Cipher cipher = InitCipherForBlock(null, 0);
                byte[] encryptedVerifier = new byte[verifier.Length];
                cipher.Update(verifier, 0, verifier.Length, encryptedVerifier);
                ver.SetEncryptedVerifier(encryptedVerifier);
                HashAlgorithm hashAlgo = ver.HashAlgorithm;
                MessageDigest hashAlg = CryptoFunctions.GetMessageDigest(hashAlgo);
                byte[] calcVerifierHash = hashAlg.Digest(verifier);
                byte[] encryptedVerifierHash = cipher.DoFinal(calcVerifierHash);
                ver.SetEncryptedVerifierHash(encryptedVerifierHash);
            }
            catch (Exception e)
            {
                throw new EncryptedDocumentException("Password Confirmation failed", e);
            }
        }

        /**
         * Initializes a cipher object for a given block index for encryption
         *
         * @param cipher may be null, otherwise the given instance is reset to the new block index
         * @param block the block index, e.g. the persist/slide id (hslf)
         * @return a new cipher object, if cipher was null, otherwise the reInitialized cipher
         * @throws GeneralSecurityException
         */
        public Cipher InitCipherForBlock(Cipher cipher, int block)
        {
            return CryptoAPIDecryptor.InitCipherForBlock(cipher, block, builder, GetSecretKey(), Cipher.ENCRYPT_MODE);
        }

        /**
         * Encrypt the Document-/SummaryInformation and other optionally streams.
         * Opposed to other crypto modes, cryptoapi is record based and can't be used
         * to stream-encrypt a whole file
         * 
         * @see <a href="http://msdn.microsoft.com/en-us/library/dd943321(v=office.12).aspx">2.3.5.4 RC4 CryptoAPI Encrypted Summary Stream</a>
         */
        public override OutputStream GetDataStream(DirectoryNode dir)
        {
            CipherByteArrayOutputStream bos = new CipherByteArrayOutputStream(this);
            byte[] buf = new byte[8];

            bos.Write(buf, 0, 8); // skip header
            String[] entryNames = {
            SummaryInformation.DEFAULT_STREAM_NAME,
            DocumentSummaryInformation.DEFAULT_STREAM_NAME
        };

            List<CryptoAPIDecryptor.StreamDescriptorEntry> descList = new List<CryptoAPIDecryptor.StreamDescriptorEntry>();

            int block = 0;
            foreach (String entryName in entryNames)
            {
                if (!dir.HasEntry(entryName)) continue;
                CryptoAPIDecryptor.StreamDescriptorEntry descEntry = new CryptoAPIDecryptor.StreamDescriptorEntry();
                descEntry.block = block;
                descEntry.streamOffset = (int)bos.Length;
                descEntry.streamName = entryName;
                descEntry.flags = CryptoAPIDecryptor.StreamDescriptorEntry.flagStream.SetValue(0, 1);
                descEntry.reserved2 = 0;

                bos.SetBlock(block);
                DocumentInputStream dis = dir.CreateDocumentInputStream(entryName);
                IOUtils.Copy(dis, bos);
                dis.Close();

                descEntry.streamSize =(int)( bos.Length - descEntry.streamOffset);
                descList.Add(descEntry);

                dir.GetEntry(entryName).Delete();

                block++;
            }

            int streamDescriptorArrayOffset = (int)bos.Length;

            bos.SetBlock(0);
            LittleEndian.PutUInt(buf, 0, descList.Count);
            bos.Write(buf, 0, 4);

            foreach (CryptoAPIDecryptor.StreamDescriptorEntry sde in descList)
            {
                LittleEndian.PutUInt(buf, 0, sde.streamOffset);
                bos.Write(buf, 0, 4);
                LittleEndian.PutUInt(buf, 0, sde.streamSize);
                bos.Write(buf, 0, 4);
                LittleEndian.PutUShort(buf, 0, sde.block);
                bos.Write(buf, 0, 2);
                LittleEndian.PutUByte(buf, 0, (short)sde.streamName.Length);
                bos.Write(buf, 0, 1);
                LittleEndian.PutUByte(buf, 0, (short)sde.flags);
                bos.Write(buf, 0, 1);
                LittleEndian.PutUInt(buf, 0, sde.reserved2);
                bos.Write(buf, 0, 4);
                byte[] nameBytes = StringUtil.GetToUnicodeLE(sde.streamName);
                bos.Write(nameBytes, 0, nameBytes.Length);
                LittleEndian.PutShort(buf, 0, (short)0); // null-termination
                bos.Write(buf, 0, 2);
            }

            int savedSize = (int)bos.Length;
            int streamDescriptorArraySize = savedSize - streamDescriptorArrayOffset;
            LittleEndian.PutUInt(buf, 0, streamDescriptorArrayOffset);
            LittleEndian.PutUInt(buf, 4, streamDescriptorArraySize);

            bos.Reset();
            bos.SetBlock(0);
            bos.Write(buf, 0, 8);
            bos.SetSize(savedSize);

            dir.CreateDocument("EncryptedSummary", new MemoryStream(bos.GetBuf(), 0, savedSize));
            DocumentSummaryInformation dsi = PropertySetFactory.NewDocumentSummaryInformation();

            try
            {
                dsi.Write(dir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            }
            catch (WritingNotSupportedException e)
            {
                throw new IOException(e.Message);
            }

            //return bos;
            throw new NotImplementedException("CipherByteArrayOutputStream should be derived from OutputStream");
        }

        protected int GetKeySizeInBytes()
        {
            return builder.GetHeader().KeySize / 8;
        }

        protected void CreateEncryptionInfoEntry(DirectoryNode dir)
        {
            DataSpaceMapUtils.AddDefaultDataSpace(dir);
            EncryptionInfo info = builder.GetEncryptionInfo();
            CryptoAPIEncryptionHeader header = builder.GetHeader();
            CryptoAPIEncryptionVerifier verifier = builder.GetVerifier();
            EncryptionRecord er = new EncryptionRecordInternal(info, header, verifier);
            DataSpaceMapUtils.CreateEncryptionEntry(dir, "EncryptionInfo", er);
        }
        private class EncryptionRecordInternal : EncryptionRecord
        {
            public EncryptionRecordInternal(EncryptionInfo info,
                StandardEncryptionHeader header, StandardEncryptionVerifier verifier)
            {
                this.info = info;
                this.header = header;
                this.verifier = verifier;
            }
            EncryptionInfo info;
            StandardEncryptionHeader header;
            StandardEncryptionVerifier verifier;
            public void Write(LittleEndianByteArrayOutputStream bos)
            {
                bos.WriteShort(info.VersionMajor);
                bos.WriteShort(info.VersionMinor);
                header.Write(bos);
                verifier.Write(bos);
            }
        }
        private class CipherByteArrayOutputStream : ByteArrayOutputStream
        {
            CryptoAPIEncryptor encryptor;
            public CipherByteArrayOutputStream(CryptoAPIEncryptor encryptor)
                : base()
            {
                this.encryptor = encryptor;
            }
            Cipher cipher;
            byte[] oneByte = { 0 };

            public CipherByteArrayOutputStream()
            {
                SetBlock(0);
            }

            public byte[] GetBuf()
            {
                return base.ToArray();
            }

            public void SetSize(long count)
            {
                base.SetLength(count);
            }

            public void SetBlock(int block)
            {
                cipher = encryptor.InitCipherForBlock(cipher, block);
            }

            public new void Write(int b)
            {
                try
                {
                    oneByte[0] = (byte)b;
                    cipher.Update(oneByte, 0, 1, oneByte, 0);
                    base.Write(oneByte);
                }
                catch (Exception e)
                {
                    throw new EncryptedDocumentException(e);
                }
            }

            public new void Write(byte[] b, int off, int len)
            {
                try
                {
                    cipher.Update(b, off, len, b, off);
                    base.Write(b, off, len);
                }
                catch (Exception e)
                {
                    throw new EncryptedDocumentException(e);
                }
            }

        }
    }

}
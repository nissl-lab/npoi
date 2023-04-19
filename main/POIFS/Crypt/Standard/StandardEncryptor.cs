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
    using System.IO;
    using System.Security.AccessControl;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.EventFileSystem;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public class StandardEncryptor : Encryptor
    {
        private StandardEncryptionInfoBuilder builder;

        protected internal StandardEncryptor(StandardEncryptionInfoBuilder builder)
        {
            this.builder = builder;
        }

        public override void ConfirmPassword(String password)
        {
            // see [MS-OFFCRYPTO] - 2.3.3 EncryptionVerifier
            //Random r = new SecureRandom();
            Random r = new Random();
            byte[] salt = new byte[16], verifier = new byte[16];
            r.NextBytes(salt);
            r.NextBytes(verifier);

            ConfirmPassword(password, null, null, salt, verifier, null);
        }


        /**
         * Fills the fields of verifier and header with the calculated hashes based
         * on the password and a random salt
         * 
         * see [MS-OFFCRYPTO] - 2.3.4.7 ECMA-376 Document Encryption Key Generation
         */
        public override void ConfirmPassword(String password, byte[] keySpec, byte[] keySalt, byte[] verifier, byte[] verifierSalt, byte[] integritySalt)
        {
            StandardEncryptionVerifier ver = builder.GetVerifier();

            ver.SetSalt(verifierSalt);
            ISecretKey secretKey = StandardDecryptor.GenerateSecretKey(password, ver, GetKeySizeInBytes());
            SetSecretKey(secretKey);
            Cipher cipher = GetCipher(secretKey, null);

            try
            {
                byte[] encryptedVerifier = cipher.DoFinal(verifier);
                MessageDigest hashAlgo = CryptoFunctions.GetMessageDigest(ver.HashAlgorithm);
                byte[] calcVerifierHash = hashAlgo.Digest(verifier);

                // 2.3.3 EncryptionVerifier ...
                // An array of bytes that Contains the encrypted form of the 
                // hash of the randomly generated Verifier value. The length of the array MUST be the size of 
                // the encryption block size multiplied by the number of blocks needed to encrypt the hash of the 
                // Verifier. If the encryption algorithm is RC4, the length MUST be 20 bytes. If the encryption 
                // algorithm is AES, the length MUST be 32 bytes. After decrypting the EncryptedVerifierHash
                // field, only the first VerifierHashSize bytes MUST be used.
                int encVerHashSize = ver.CipherAlgorithm.encryptedVerifierHashLength;
                byte[] encryptedVerifierHash = cipher.DoFinal(Arrays.CopyOf(calcVerifierHash, encVerHashSize));

                ver.SetEncryptedVerifier(encryptedVerifier);
                ver.SetEncryptedVerifierHash(encryptedVerifierHash);
            }
            catch (Exception e)
            {
                throw new EncryptedDocumentException("Password Confirmation failed", e);
            }

        }

        private Cipher GetCipher(ISecretKey key, string padding)
        {
            EncryptionVerifier ver = builder.GetVerifier();
            return CryptoFunctions.GetCipher(key, ver.CipherAlgorithm, ver.ChainingMode, null, Cipher.ENCRYPT_MODE, padding);
        }

        public override OutputStream GetDataStream(DirectoryNode dir)
        {
            CreateEncryptionInfoEntry(dir);
            DataSpaceMapUtils.AddDefaultDataSpace(dir);
            Stream countStream = new StandardCipherOutputStream(dir, this);
            //return countStream;
            throw new NotImplementedException("StandardCipherOutputStream should be derived from OutputStream");
        }

        protected class StandardCipherOutputStream : ByteArrayOutputStream, POIFSWriterListener
        {
            private StandardEncryptor encryptor;
            protected long countBytes;
            protected FileInfo fileOut;
            protected DirectoryNode dir;
            ByteArrayOutputStream out1;
            FileStream rawStream;// maybe has memory leak problem.

            protected internal StandardCipherOutputStream(DirectoryNode dir, StandardEncryptor encryptor)
            {
                this.encryptor = encryptor;
                this.dir = dir;
                fileOut = TempFile.CreateTempFile("encrypted_package", "crypt");
                rawStream = new FileStream(fileOut.FullName, FileMode.Open, FileAccess.ReadWrite); // fileOut.Create();

                // although not documented, we need the same padding as with agile encryption
                // and instead of calculating the missing bytes for the block size ourselves
                // we leave it up to the CipherOutputStream, which generates/saves them on close()
                // ... we can't use "NoPAdding" here
                //
                // see also [MS-OFFCRYPT] - 2.3.4.15
                // The data block MUST be padded to the next integral multiple of the
                // KeyData.blockSize value. Any pAdding bytes can be used. Note that the StreamSize
                // field of the EncryptedPackage field specifies the number of bytes of 
                // unencrypted data as specified in section 2.3.4.4.
                CipherOutputStream cryptStream = new CipherOutputStream(rawStream, 
                    encryptor.GetCipher(encryptor.GetSecretKey(), "PKCS5Padding"));

                this.out1 = cryptStream;
            }


            public override void Write(byte[] b, int off, int len)
            {
                out1.Write(b, off, len);
                countBytes += len;
            }


            public override void Write(int b)
            {
                out1.Write(b);
                countBytes++;
            }

            public override void Close()
            {
                // the CipherOutputStream Adds the pAdding bytes on close()
                base.Close();
                WriteToPOIFS();
                //rawStream.Close();
                //rawStream = null;
            }

            void WriteToPOIFS()
            {
                int oleStreamSize = (int)(fileOut.Length + LittleEndianConsts.LONG_SIZE);
                dir.CreateDocument(DEFAULT_POIFS_ENTRY, oleStreamSize, this);
                // TODO: any properties???
            }

            public void ProcessPOIFSWriterEvent(POIFSWriterEvent event1)
            {
                try
                {
                    LittleEndianOutputStream leos = new LittleEndianOutputStream(event1.Stream);

                    // StreamSize (8 bytes): An unsigned integer that specifies the number of bytes used by data 
                    // encrypted within the EncryptedData field, not including the size of the StreamSize field. 
                    // Note that the actual size of the \EncryptedPackage stream (1) can be larger than this 
                    // value, depending on the block size of the chosen encryption algorithm
                    leos.WriteLong(countBytes);
                    long rawPos = rawStream.Position;
                    //FileStream fis = new FileStream(fileOut.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    IOUtils.Copy(rawStream, leos.out1);
                    //fis.Close();
                    rawStream.Position = rawPos;
                    //File.Delete(fileOut.FullName + ".copy");
                    fileOut.Delete();
                    

                    leos.Close();
                }
                catch (IOException e)
                {
                    throw new EncryptedDocumentException(e);
                }
            }
        }

        protected int GetKeySizeInBytes()
        {
            return builder.GetHeader().KeySize / 8;
        }

        protected internal void CreateEncryptionInfoEntry(DirectoryNode dir)
        {
            EncryptionInfo info = builder.GetEncryptionInfo();
            StandardEncryptionHeader header = builder.GetHeader();
            StandardEncryptionVerifier verifier = builder.GetVerifier();

            EncryptionRecord er = new EncryptionRecordInternal(info, header, verifier);


            DataSpaceMapUtils.CreateEncryptionEntry(dir, "EncryptionInfo", er);

            // TODO: any properties???
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
                bos.WriteInt(info.EncryptionFlags);
                header.Write(bos);
                verifier.Write(bos);
            }
        }

    }

    internal class CipherOutputStream : ByteArrayOutputStream
    {
        private byte[] ibuffer = new byte[1];
        private byte[] obuffer;
        private bool closed = false;
        private FileStream output;
        private Cipher cipher;

        public CipherOutputStream(FileStream rawStream, Cipher cipher)
        {
            this.output = rawStream;
            this.cipher = cipher;
        }
        protected CipherOutputStream(FileStream rawStream)
        {
            this.output = rawStream;
            this.cipher = new NullCipher();
        }
        public override void Write(int paramInt)
        {
            this.ibuffer[0] = ((byte)paramInt);
            this.obuffer = this.cipher.Update(this.ibuffer, 0, 1);
            if (this.obuffer != null)
            {
                this.output.Write(this.obuffer, 0, obuffer.Length);
                this.obuffer = null;
            }
        }
        public override void Write(byte[] b)
        {
            Write(b, 0, b.Length);
        }

        public override void Write(byte[] b, int off, int len)
        {
            this.obuffer = this.cipher.Update(b, off, len);
            if (this.obuffer != null)
            {
                this.output.Write(this.obuffer, 0, obuffer.Length);
                this.output.Flush();
                this.obuffer = null;
            }
        }

        public override void Flush()
        {
            if (this.obuffer != null)
            {
                this.output.Write(this.obuffer, 0, obuffer.Length);
                this.obuffer = null;
            }
            this.output.Flush();
        }

        public override void Close()
        {
            if (this.closed)
            {
                return;
            }
            this.closed = true;
            try
            {
                this.obuffer = this.cipher.DoFinal();
            }
            catch
            {
                this.obuffer = null;
            }
            try
            {
                Flush();
            }
            catch (IOException) { }
            this.Close();
        }
    }
}
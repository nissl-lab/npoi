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
    using Org.BouncyCastle.Crypto;

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

        protected internal static ISecretKey GenerateSecretKey(string password, EncryptionVerifier ver, int keySize) {
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

            ISecretKey skey = new SecretKeySpec(key, ver.CipherAlgorithm.jceId);
            return skey;
        }

        protected static byte[] FillAndXor(byte[] hash, byte FillByte) {
            byte[] buff = new byte[64];
            Arrays.Fill(buff, FillByte);

            for (int i = 0; i < hash.Length; i++) {
                buff[i] = (byte)(buff[i] ^ hash[i]);
            }

            MessageDigest sha1 = CryptoFunctions.GetMessageDigest(HashAlgorithm.sha1);
            return sha1.Digest(buff);
        }

        private Cipher GetCipher(ISecretKey key) {
            EncryptionHeader em = builder.GetHeader();
            ChainingMode cm = em.ChainingMode;
            Debug.Assert(cm == ChainingMode.ecb);

            return CryptoFunctions.GetCipher(key, em.CipherAlgorithm, cm, null, Cipher.DECRYPT_MODE);
        }

        public override InputStream GetDataStream(DirectoryNode dir) {
            DocumentInputStream dis = dir.CreateDocumentInputStream(Encryptor.DEFAULT_POIFS_ENTRY);

            _length = dis.ReadLong();
            if (GetSecretKey() == null)
            {
                VerifyPassword(null);
            }
            // limit wrong calculated ole entries - (bug #57080)
            // standard encryption always uses aes encoding, so blockSize is always 16 
            // http://stackoverflow.com/questions/3283787/size-of-data-after-aes-encryption
            int blockSize = builder.GetHeader().CipherAlgorithm.blockSize;
            long cipherLen = (_length / blockSize + 1) * blockSize;
            Cipher cipher = GetCipher(GetSecretKey());


            ByteArrayInputStream boundedDis = new BoundedInputStream(dis, cipherLen);
            return new BoundedInputStream(new CipherInputStream(boundedDis, cipher), _length);
        }

        /**
         * @return the length of the stream returned by {@link #getDataStream(DirectoryNode)}
         */
        public override long GetLength() {
            if (_length == -1) throw new InvalidOperationException("Decryptor.DataStream was not called");
            return _length;
        }
    }

    public class CipherInputStream : ByteArrayInputStream
    {
        private Cipher cipher;
        private ByteArrayInputStream input;
        private byte[] ibuffer = new byte['?'];
        private bool done = false;
        private byte[] obuffer;
        private int ostart = 0;
        private int ofinish = 0;
        private bool closed = false;
        public CipherInputStream(ByteArrayInputStream paramInputStream, Cipher paramCipher)
        //: base(paramInputStream)
        {
            this.input = paramInputStream;
            this.cipher = paramCipher;
        }
        protected CipherInputStream(ByteArrayInputStream paramInputStream)
        //   : base(paramInputStream)
        {
            this.input = paramInputStream;
            this.cipher = new NullCipher();
        }
        private int getMoreData()
        {
            if (this.done)
            {
                return -1;
            }
            int i = this.input.Read(this.ibuffer, 0, this.ibuffer.Length);
            if (i == -1)
            {
                this.done = true;
                try
                {
                    this.obuffer = this.cipher.DoFinal();
                }
                catch (Exception ex)
                {
                    this.obuffer = null;
                    throw new IOException(ex.Message);
                }
                if (this.obuffer == null)
                {
                    return -1;
                }
                this.ostart = 0;
                this.ofinish = this.obuffer.Length;
                return this.ofinish;
            }
            try
            {
                this.obuffer = this.cipher.Update(this.ibuffer, 0, i);
            }
            catch (Exception ex)
            {
                this.obuffer = null;
                throw ex;
            }
            this.ostart = 0;
            if (this.obuffer == null)
            {
                this.ofinish = 0;
            }
            else
            {
                this.ofinish = this.obuffer.Length;
            }
            return this.ofinish;
        }
        public override int Read()
        {
            if (this.ostart >= this.ofinish)
            {
                int i = 0;
                while (i == 0)
                {
                    i = getMoreData();
                }
                if (i == -1)
                {
                    return -1;
                }
            }
            return this.obuffer[(this.ostart++)] & 0xFF;
        }
        public int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
        }

        public override int Read(byte[] b, int off, int len)
        {
            int i;
            if (this.ostart >= this.ofinish)
            {
                i = 0;
                while (i == 0)
                {
                    i = getMoreData();
                }
                if (i == -1)
                {
                    return -1;
                }
            }
            if (len <= 0)
            {
                return 0;
            }
            i = this.ofinish - this.ostart;
            if (len < i)
            {
                i = len;
            }
            if (b != null)
            {
                Array.Copy(this.obuffer, this.ostart, b, off, i);
            }
            this.ostart += i;
            return i;
        }

        public long Skip(long paramLong)
        {
            int i = this.ofinish - this.ostart;
            if (paramLong > i)
            {
                paramLong = i;
            }
            if (paramLong < 0L)
            {
                return 0L;
            }
            this.ostart = ((int)(this.ostart + paramLong));
            return paramLong;
        }

        public override int Available()
        {
            return this.ofinish - this.ostart;
        }

        public override void Close()
        {
            if (this.closed)
            {
                return;
            }
            this.closed = true;
            this.input.Close();
            if (!this.done)
            {
                try
                {
                    this.cipher.DoFinal();
                }
                catch (Exception)
                {
                }
            }
            this.ostart = 0;
            this.ofinish = 0;
        }


        public bool MarkSupported()
        {
            return false;
        }
    }

    public class NullBufferedCipher : IBufferedCipher
    {
        public string AlgorithmName
        {
            get { return "Null"; }
        }

        public byte[] DoFinal()
        {
            return new byte[0];
        }

        public byte[] DoFinal(byte[] input)
        {
            return new byte[0];
        }

        public byte[] DoFinal(byte[] input, int inOff, int length)
        {
            return new byte[0];
        }

        public int DoFinal(byte[] output, int outOff)
        {
            return 0;
        }

        public int DoFinal(byte[] input, byte[] output, int outOff)
        {
            return 0;
        }

        public int DoFinal(byte[] input, int inOff, int length, byte[] output, int outOff)
        {
            return 0;
        }

        public int GetBlockSize()
        {
            return 0;
        }

        public int GetOutputSize(int inputLen)
        {
            return 0;
        }

        public int GetUpdateOutputSize(int inputLen)
        {
            return 0;
        }

        public void Init(bool forEncryption, ICipherParameters parameters)
        {
            
        }

        public byte[] ProcessByte(byte input)
        {
            return new byte[0];
        }

        public int ProcessByte(byte input, byte[] output, int outOff)
        {
            return 0;
        }

        public byte[] ProcessBytes(byte[] input)
        {
            return new byte[0];
        }

        public byte[] ProcessBytes(byte[] input, int inOff, int length)
        {
            return new byte[0];
        }

        public int ProcessBytes(byte[] input, byte[] output, int outOff)
        {
            return 0;
        }

        public int ProcessBytes(byte[] input, int inOff, int length, byte[] output, int outOff)
        {
            return 0;
        }

        public void Reset()
        {
            
        }
    }
    public class NullCipher : Cipher
    {
        public NullCipher()
        {
            cipherImpl = new NullBufferedCipher();
        }
    }
}
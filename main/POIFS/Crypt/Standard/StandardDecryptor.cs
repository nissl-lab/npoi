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

            // Pass expected plain length so CipherInputStream can trim padding
            return new CipherInputStream(boundedDis, cipher, _length);
        }

        /**
         * @return the length of the stream returned by {@link #getDataStream(DirectoryNode)}
         */
        public override long GetLength() {
            if (_length == -1) throw new InvalidOperationException("Decryptor.DataStream was not called");
            return _length;
        }
    }

    public class CipherInputStream : FilterInputStream
    {
        // Decrypts the encrypted package stream produced by StandardEncryptor.CipherOutputStream.
        // Handles block cipher with PKCS5/PKCS7 padding and trims padding bytes so that only the
        // original unencrypted length (passed from StandardDecryptor) is exposed.
        private readonly Cipher cipher;
        private readonly ByteArrayInputStream encryptedInput; // bounded encrypted stream
        private readonly long expectedPlainLength;            // original unencrypted length (StreamSize)
        private bool finalized;                               // cipher.DoFinal executed
        private bool closed;
        private long totalPlainProduced;
        private byte[] encBuffer = new byte[4096];
        private byte[] plainBuffer = Array.Empty<byte>();
        private int plainOffset;
        private int plainCount;

        public CipherInputStream(ByteArrayInputStream encryptedInput, Cipher cipher, long expectedPlainLength)
            : base(encryptedInput)
        {
            this.encryptedInput = encryptedInput;
            this.cipher = cipher;
            this.expectedPlainLength = expectedPlainLength;
        }

        // Legacy protected ctor kept (NullCipher fallback)
        protected CipherInputStream(ByteArrayInputStream paramInputStream)
            : base(paramInputStream)
        {
            this.encryptedInput = paramInputStream;
            this.cipher = new NullCipher();
            this.expectedPlainLength = 0;
        }

        private bool EnsurePlainData()
        {
            if(plainOffset < plainCount)
                return true; // still have plaintext

            // Reset consumption markers
            plainOffset = 0;
            plainCount = 0;

            if(finalized)
                return false;

            int read = encryptedInput.Read(encBuffer, 0, encBuffer.Length);
            if(read > 0)
            {
                var block = cipher.Update(encBuffer, 0, read);
                if(block != null && block.Length > 0)
                {
                    SetPlainBuffer(block, false);
                    return plainCount > 0;
                }
                // No plaintext produced yet, keep looping
                return EnsurePlainData();
            }

            // End of encrypted stream: finalize
            byte[] finalPlain;
            try
            {
                finalPlain = cipher.DoFinal();
            }
            catch
            {
                finalPlain = Array.Empty<byte>();
            }
            finalized = true;
            if(finalPlain.Length > 0)
            {
                SetPlainBuffer(finalPlain, true);
                return plainCount > 0;
            }
            return false;
        }

        private void SetPlainBuffer(byte[] data, bool isFinal)
        {
            // Trim any excess beyond expectedPlainLength (remove padding visibility)
            long remaining = expectedPlainLength - totalPlainProduced;
            if(remaining <= 0)
            {
                // Already satisfied length; ignore further (padding)
                plainBuffer = Array.Empty<byte>();
                plainCount = 0;
                plainOffset = 0;
                return;
            }

            int usable = data.Length;
            if(usable > remaining)
                usable = (int) remaining;

            if(usable == data.Length)
            {
                plainBuffer = data;
            }
            else
            {
                // Copy only the usable portion (discard padding tail)
                plainBuffer = new byte[usable];
                Array.Copy(data, 0, plainBuffer, 0, usable);
            }
            plainCount = plainBuffer.Length;
            plainOffset = 0;
            totalPlainProduced += plainCount;

            if(isFinal)
            {
                // After final block, mark finalized even if we truncated
                finalized = true;
            }
        }

        public override int Read()
        {
            if(!EnsurePlainData())
                return -1;
            int val = plainBuffer[plainOffset++] & 0xFF;
            return val;
        }

        public int Read(byte[] b) => Read(b, 0, b.Length);

        public override int Read(byte[] b, int off, int len)
        {
            if(off < 0 || len < 0 || (off + len) > b.Length)
                throw new ArgumentOutOfRangeException();

            if(len == 0)
                return 0;
            if(!EnsurePlainData())
                return -1;

            int available = plainCount - plainOffset;
            if(available <= 0 && !EnsurePlainData())
                return -1;
            available = plainCount - plainOffset;

            int toCopy = available < len ? available : len;
            Array.Copy(plainBuffer, plainOffset, b, off, toCopy);
            plainOffset += toCopy;
            return toCopy;
        }

        public long Skip(long n)
        {
            if(n <= 0)
                return 0;
            long skipped = 0;
            while(skipped < n)
            {
                if(!EnsurePlainData())
                    break;
                int available = plainCount - plainOffset;
                if(available <= 0)
                    break;
                int step = (int)Math.Min(available, n - skipped);
                plainOffset += step;
                skipped += step;
            }
            return skipped;
        }

        public override int Available()
        {
            return (plainCount - plainOffset);
        }

        public override void Close()
        {
            if(closed)
                return;
            closed = true;
            // Drain remaining to finalize cipher (optional)
            if(!finalized)
            {
                try
                {
                    while(EnsurePlainData())
                    {
                        plainOffset = plainCount; // discard
                    }
                }
                catch { }
            }
            encryptedInput.Close();
            plainBuffer = Array.Empty<byte>();
            plainCount = 0;
            plainOffset = 0;
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
            return [];
        }

        public byte[] DoFinal(byte[] input)
        {
            return [];
        }

        public byte[] DoFinal(byte[] input, int inOff, int length)
        {
            return [];
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
            return [];
        }

        public int ProcessByte(byte input, byte[] output, int outOff)
        {
            return 0;
        }

        public byte[] ProcessBytes(byte[] input)
        {
            return [];
        }

        public byte[] ProcessBytes(byte[] input, int inOff, int length)
        {
            return [];
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

#if NET6_0_OR_GREATER
        public int ProcessByte(byte input, Span<byte> output)
        {
            return 0;
        }

        public int ProcessBytes(ReadOnlySpan<byte> input, Span<byte> output)
        {
            return 0;
        }

        public int DoFinal(Span<byte> output)
        {
            return 0;
        }

        public int DoFinal(ReadOnlySpan<byte> input, Span<byte> output)
        {
            return 0;
        }
#endif
    }
    public class NullCipher : Cipher
    {
        public NullCipher()
        {
            cipherImpl = new NullBufferedCipher();
        }
    }
}
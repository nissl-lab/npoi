namespace NPOI.POIFS.Crypt.XOR
{
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using System;
    using System.IO;

    public class XORDecryptor : Decryptor
    {
        private long _length = -1L;
        private int chunkSize = 512;

        private sealed class XORCipherInputStream : ChunkedCipherInputStream
        {
            protected override Cipher InitCipherForBlock(Cipher existing, int block)
            {
                return XORDecryptor.InitCipherForBlock(existing, block, decryptor.GetSecretKey(), Cipher.DECRYPT_MODE);
            }

            public XORCipherInputStream(InputStream stream, long size, int chunkSize, XORDecryptor decryptor)
                : base(stream, size, chunkSize, decryptor)
            {
            }
        }

        public XORDecryptor(IEncryptionInfoBuilder builder) : base(builder)
        {
        }

        public override bool VerifyPassword(string password)
        {
            int keyComp = CryptoFunctions.CreateXorKey1(password);
            int verifierComp = CryptoFunctions.CreateXorVerifier1(password);
            
            XOREncryptionVerifier ver = (XOREncryptionVerifier)builder.GetVerifier();
            int encKey = ver.GetEncryptedKey();
            int encVerifier = ver.GetEncryptedVerifier();
            
            if (keyComp == encKey && verifierComp == encVerifier)
            {
                byte[] xorArray = CryptoFunctions.CreateXorArray1(password);
                SetSecretKey(new SecretKeySpec(xorArray, "XOR"));
                return true;
            }
            
            return false;
        }

        public static Cipher InitCipherForBlock(Cipher cipher, int block, ISecretKey skey, int encryptMode)
        {
            // XOR encryption doesn't use traditional ciphers, but we need to return something
            // The actual XOR logic is handled in the stream processing
            if (cipher == null)
            {
                cipher = new XORCipher(skey, encryptMode);
            }
            return cipher;
        }

        public override InputStream GetDataStream(DirectoryNode dir)
        {
            DocumentInputStream dis = dir.CreateDocumentInputStream(DEFAULT_POIFS_ENTRY);
            _length = dis.ReadLong();
            XORCipherInputStream cipherStream = new XORCipherInputStream(dis, _length, chunkSize, this);
            return cipherStream;
        }

        public InputStream GetDataStream(InputStream stream, int size, int initialPos)
        {
            return new XORCipherInputStream(stream, initialPos, chunkSize, this);
        }

        public override long GetLength()
        {
            if (_length == -1L)
            {
                throw new InvalidOperationException("Decryptor.DataStream was not called");
            }

            return _length;
        }

        // Simple XOR cipher implementation for compatibility
        private class XORCipher : Cipher
        {
            private ISecretKey secretKey;
            private int mode;

            public XORCipher(ISecretKey secretKey, int mode)
            {
                this.secretKey = secretKey;
                this.mode = mode;
            }

            public void Init(int opmode, ISecretKey key)
            {
                this.mode = opmode;
                this.secretKey = key;
            }

            public byte[] Update(byte[] input, int inputOffset, int inputLen, byte[] output)
            {
                byte[] xorArray = secretKey.GetEncoded();
                for (int i = 0; i < inputLen; i++)
                {
                    output[i] = (byte)(input[inputOffset + i] ^ xorArray[i % xorArray.Length]);
                }
                return output;
            }

            public byte[] DoFinal(byte[] input)
            {
                byte[] output = new byte[input.Length];
                Update(input, 0, input.Length, output);
                return output;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.POIFS.Crypt
{
    public class Cipher
    {
        internal static int DECRYPT_MODE;
        internal static int ENCRYPT_MODE;

        public byte[] DoFinal(byte[] block)
        {
            throw new NotImplementedException();
        }
        public byte[] DoFinal(byte[] block, int v, int posInChunk)
        {
            throw new NotImplementedException();
        }

        internal int DoFinal(byte[] chunk1, int v, int posInChunk, byte[] chunk2)
        {
            throw new NotImplementedException();
        }

        internal static Cipher GetInstance(string jceId)
        {
            throw new NotImplementedException();
        }

        internal void Init(int cipherMode, IKey key, AlgorithmParameterSpec aps)
        {
            throw new NotImplementedException();
        }

        internal static Cipher GetInstance(string v1, string v2)
        {
            throw new NotImplementedException();
        }

        internal static int GetMaxAllowedKeyLength(string jceId)
        {
            throw new NotImplementedException();
        }

        internal void Init(int cipherMode, IKey key)
        {
            throw new NotImplementedException();
        }

        internal void Update(byte[] encryptedVerifier, int v, int length, byte[] verifier)
        {
            throw new NotImplementedException();
        }

        internal void Update(byte[] b1, int off1, int readLen, byte[] b2, int off2)
        {
            throw new NotImplementedException();
        }
    }
}

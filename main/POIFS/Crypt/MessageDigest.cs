using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.POIFS.Crypt
{
    public class MessageDigest
    {
        public byte[] Digest(byte[] data, int v)
        {
            throw new NotImplementedException();
        }

        internal static MessageDigest GetInstance(string jceId, string v)
        {
            throw new NotImplementedException();
        }

        internal static MessageDigest GetInstance(string jceId)
        {
            throw new NotImplementedException();
        }

        internal void Update(byte[] passwordHash)
        {
            throw new NotImplementedException();
        }

        internal void Reset()
        {
            throw new NotImplementedException();
        }

        internal void Digest(byte[] hash, int v, int length)
        {
            throw new NotImplementedException();
        }


        internal Array Digest()
        {
            throw new NotImplementedException();
        }

        internal void Update(byte[] hash, int v1, int v2)
        {
            throw new NotImplementedException();
        }

        public byte[] Digest(byte[] verifier)
        {
            throw new NotImplementedException();
        }
    }
}

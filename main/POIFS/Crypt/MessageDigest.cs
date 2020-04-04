using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Security;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.POIFS.Crypt
{
    public class MessageDigest
    {
        private IDigest digestImpl;
        public MessageDigest()
        {

        }

        internal static MessageDigest GetInstance(string jceId, string v)
        {
            return GetInstance(jceId);
        }

        internal static MessageDigest GetInstance(string jceId)
        {
            var md = new MessageDigest()
            {
                digestImpl = DigestUtilities.GetDigest(jceId)
            };
            
            return md;
        }

        internal void Update(byte[] passwordHash)
        {
            Update(passwordHash, 0, passwordHash.Length);
        }

        internal void Reset()
        {
            digestImpl.Reset();
        }

        internal int Digest(byte[] buf, int offset, int len)
        {
            if (buf == null)
            {
                throw new ArgumentNullException("No output buffer given");
            }
            if (buf.Length - offset < len)
            {
                throw new ArgumentOutOfRangeException
                    ("Output buffer too small for specified offset and length");
            }
            byte[] digest = Digest();
            if (len < digest.Length)
                throw new Exception("partial digests not returned");
            if (buf.Length - offset < digest.Length)
                throw new Exception("insufficient space in the output "
                                          + "buffer to store the digest");
            Array.Copy(digest, 0, buf, offset, digest.Length);
            return digest.Length;
        }


        internal byte[] Digest()
        {
            byte[] resBuf = new byte[digestImpl.GetDigestSize()];
            digestImpl.DoFinal(resBuf, 0);
            return resBuf;
        }

        internal void Update(byte[] hash, int v1, int v2)
        {
            digestImpl.BlockUpdate(hash, v1, v2);
        }

        public byte[] Digest(byte[] input)
        {
            byte[] resBuf = new byte[digestImpl.GetDigestSize()];
            digestImpl.BlockUpdate(input, 0, input.Length);
            digestImpl.DoFinal(resBuf, 0);
            return resBuf;
        }
    }
}

using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Security;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.POIFS.Crypt
{
    public class Cipher
    {
        public const int DECRYPT_MODE = 2;
        public const int ENCRYPT_MODE = 1;
        public const int WRAP_MODE = 3;
        public const int UNWRAP_MODE = 4;

        public const int PUBLIC_KEY = 1;
        public const int PRIVATE_KEY = 2;
        public const int SECRET_KEY = 3;

        protected IBufferedCipher cipherImpl;

        public static Cipher GetInstance(string transformation)
        {
            Cipher cihpher = new Cipher()
            {
                cipherImpl = CipherUtilities.GetCipher(transformation)
            };

            return cihpher;
        }

        public static Cipher GetInstance(string transformation, string provider)
        {
            return GetInstance(transformation);
        }

        public void Init(int cipherMode, IKey key, AlgorithmParameterSpec aps)
        {
            ICipherParameters cp;
            if (aps is RC2ParameterSpec)
            {
                cp = new RC2Parameters(key.GetEncoded(), (aps as RC2ParameterSpec).GetEffectiveKeyBits());
            }
            else if (aps is IvParameterSpec)
            {
                cp = new KeyParameter(key.GetEncoded());
                cp = new ParametersWithIV(cp, (aps as IvParameterSpec).GetIV());
            }
            else
            {
                throw new NotImplementedException();
            }
            cipherImpl.Init(cipherMode == ENCRYPT_MODE, cp);
        }

        public void Init(int cipherMode, IKey key)
        {
            ICipherParameters cp = new RC2Parameters(key.GetEncoded());
            cipherImpl.Init(cipherMode == ENCRYPT_MODE, cp);
        }
        public void Init(int cipherMode, ICipherParameters cipherParameters)
        {
            cipherImpl.Init(cipherMode == ENCRYPT_MODE, cipherParameters);
        }

        public static int GetMaxAllowedKeyLength(string jceId)
        {
            /*
             * rc4/RC4/128
             * aes128/AES/128
             * aes192/AES/128
             * aes256/AES/128
             * rc2/RC2/128
             * des/DES/64
             * des3/DESede/2147483647
             * des3_112/DESede/2147483647
             * rsa/RSA/2147483647
             */
            switch (jceId)
            {
                case "RC2": return 128;
                case "RC4": return 128;
                case "DES": return 64;
                case "AES": return 128;
                case "DESede": return 2147483647;
                case "RSA": return 2147483647;
                default:
                    throw new NotImplementedException();
            }
        }

        public int Update(byte[] input, int inputOffset, int inputLen, byte[] output)
        {
            if ((input == null) || (inputOffset < 0) || (inputLen > input.Length - inputOffset) || (inputLen < 0))
            {
                throw new ArgumentException("Bad arguments");
            }
            return cipherImpl.ProcessBytes(input, inputOffset, inputLen, output, 0);
        }

        public int Update(byte[] input, int inputOffset, int inputLen, byte[] output, int outputOffset)
        {
            if ((input == null) || (inputOffset < 0) || (inputLen > input.Length - inputOffset) || (inputLen < 0) || (outputOffset < 0))
            {
                throw new ArgumentException("Bad arguments");
            }
            return cipherImpl.ProcessBytes(input, inputOffset, inputLen, output, outputOffset);
        }

        public byte[] Update(byte[] ibuffer, int inOff, int length)
        {
            return cipherImpl.ProcessBytes(ibuffer, inOff, length);
        }

        public byte[] DoFinal(byte[] block)
        {
            return cipherImpl.DoFinal(block);
        }

        //public byte[] DoFinal(byte[] block, int v, int posInChunk)
        //{
        //    return cipherImpl.DoFinal(block, v, posInChunk);
        //}

        public int DoFinal(byte[] input, int inputOffset, int inputLen, byte[] output)
        {
            return cipherImpl.DoFinal(input, inputOffset, inputLen, output, 0);
        }

        public byte[] DoFinal()
        {
            return cipherImpl.DoFinal();
        }
    }
}

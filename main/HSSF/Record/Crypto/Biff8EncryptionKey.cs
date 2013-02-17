namespace NPOI.HSSF.Record.Crypto
{
    using System;
    using System.IO;

    using NPOI.Util;

    
    using System.Security.Cryptography;

    public class Biff8EncryptionKey
    {
        // these two constants coincidentally have the same value
        private const int KEY_DIGEST_LENGTH = 5;
        private const int PASSWORD_HASH_NUMBER_OF_BYTES_USED = 5;

        private byte[] _keyDigest;

        /**
         * Create using the default password and a specified docId
         * @param docId 16 bytes
         */
        public static Biff8EncryptionKey Create(byte[] docId)
        {
            return new Biff8EncryptionKey(CreateKeyDigest("VelvetSweatshop", docId));
        }
        public static Biff8EncryptionKey Create(String password, byte[] docIdData)
        {
            return new Biff8EncryptionKey(CreateKeyDigest(password, docIdData));
        }

        internal Biff8EncryptionKey(byte[] keyDigest)
        {
            if (keyDigest.Length != KEY_DIGEST_LENGTH)
            {
                throw new ArgumentException("Expected 5 byte key digest, but got " + HexDump.ToHex(keyDigest));
            }
            _keyDigest = keyDigest;
        }

        internal static byte[] CreateKeyDigest(String password, byte[] docIdData)
        {
            Check16Bytes(docIdData, "docId");
            int nChars = Math.Min(password.Length, 16);
            byte[] passwordData = new byte[nChars * 2];
            for (int i = 0; i < nChars; i++)
            {
                char ch = password[i];
                passwordData[i * 2 + 0] = (byte)((ch << 0) & 0xFF);
                passwordData[i * 2 + 1] = (byte)((ch << 8) & 0xFF);
            }

            byte[] kd;
            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                byte[] passwordHash = md5.ComputeHash(passwordData);

                md5.Initialize();

                byte[] data = new byte[PASSWORD_HASH_NUMBER_OF_BYTES_USED * 16 + docIdData.Length * 16];

                int offset = 0;
                for (int i = 0; i < 16; i++)
                {
                    Array.Copy(passwordHash, 0, data, offset, PASSWORD_HASH_NUMBER_OF_BYTES_USED);
                    offset += PASSWORD_HASH_NUMBER_OF_BYTES_USED;// passwordHash.Length;
                    Array.Copy(docIdData, 0, data, offset, docIdData.Length);
                    offset += docIdData.Length;
                }
                kd = md5.ComputeHash(data);
                byte[] result = new byte[KEY_DIGEST_LENGTH];
                Array.Copy(kd, 0, result, 0, KEY_DIGEST_LENGTH);
                md5.Clear();

                return result;
            }
        }

        /**
         * @return <c>true</c> if the keyDigest is compatible with the specified saltData and saltHash
         */
        public bool Validate(byte[] saltData, byte[] saltHash)
        {
            Check16Bytes(saltData, "saltData");
            Check16Bytes(saltHash, "saltHash");

            // validation uses the RC4 for block zero
            RC4 rc4 = CreateRC4(0);
            byte[] saltDataPrime = new byte[saltData.Length];
            Array.Copy(saltData, saltDataPrime, saltData.Length);
            rc4.Encrypt(saltDataPrime);

            byte[] saltHashPrime = new byte[saltHash.Length];
            Array.Copy(saltHash, saltHashPrime, saltHash.Length);
            rc4.Encrypt(saltHashPrime);

            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                byte[] finalSaltResult = md5.ComputeHash(saltDataPrime);

                //if (false)
                //{ // set true to see a valid saltHash value
                //    byte[] saltHashThatWouldWork = xor(saltHash, xor(saltHashPrime, finalSaltResult));
                //    Console.WriteLine(HexDump.ToHex(saltHashThatWouldWork));
                //}

                return Arrays.Equals(saltHashPrime, finalSaltResult);
            }
        }

        private static byte[] xor(byte[] a, byte[] b)
        {
            byte[] c = new byte[a.Length];
            for (int i = 0; i < c.Length; i++)
            {
                c[i] = (byte)(a[i] ^ b[i]);
            }
            return c;
        }
        private static void Check16Bytes(byte[] data, String argName)
        {
            if (data.Length != 16)
            {
                throw new ArgumentException("Expected 16 byte " + argName + ", but got " + HexDump.ToHex(data));
            }
        }

        //private static ConcatBytes()

        /**
         * The {@link RC4} instance needs to be Changed every 1024 bytes.
         * @param keyBlockNo used to seed the newly Created {@link RC4}
         */
        internal RC4 CreateRC4(int keyBlockNo)
        {
            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                using (MemoryStream baos = new MemoryStream(4))
                {
                    new LittleEndianOutputStream(baos).WriteInt(keyBlockNo);
                    byte[] baosToArray = baos.ToArray();
                    byte[] data = new byte[baosToArray.Length + _keyDigest.Length];
                    Array.Copy(_keyDigest, 0, data, 0, _keyDigest.Length);
                    Array.Copy(baosToArray, 0, data, _keyDigest.Length, baosToArray.Length);

                    byte[] digest = md5.ComputeHash(data);
                    return new RC4(digest);
                }
            }
        }


        /**
         * Stores the BIFF8 encryption/decryption password for the current thread.  This has been done
         * using a {@link ThreadLocal} in order to avoid further overloading the various public APIs
         * (e.g. {@link HSSFWorkbook}) that need this functionality.
         */
        [ThreadStatic]
        private static String _userPasswordTLS = null;

        /**
         * @return the BIFF8 encryption/decryption password for the current thread.
         * <code>null</code> if it is currently unSet.
         */
        public static String CurrentUserPassword
        {
            get
            {
                return _userPasswordTLS;
            }
            set 
            {
                _userPasswordTLS = value;
            }
        }
    }

}

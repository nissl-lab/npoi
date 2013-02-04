/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

//http://stackoverflow.com/questions/6843698/calculating-sha-1-hashes-in-java-and-c-sharp  Leon
using System;
using NPOI.POIFS.FileSystem;
using System.Security.Cryptography;
using System.IO;

using NPOI;
using NPOI.Util;


namespace NPOI.POIFS.Crypt
{
    public class AgileDecryptor : Decryptor
    {
        private EncryptionInfo _info;
        //private SecretKey _secretKey; 
        private byte[] _secretKey;
        private long _length = -1;

        private static byte[] kVerifierInputBlock =
            new byte[] { (byte)0xfe, (byte)0xa7, (byte)0xd2, (byte)0x76,
                         (byte)0x3b, (byte)0x4b, (byte)0x9e, (byte)0x79 };

        private static byte[]  kHashedVerifierBlock =
                new byte[] { (byte)0xd7, (byte)0xaa, (byte)0x0f, (byte)0x6d,
                             (byte)0x30, (byte)0x61, (byte)0x34, (byte)0x4e };


        private static byte[] kCryptoKeyBlock =
                new byte[] { (byte)0x14, (byte)0x6e, (byte)0x0b, (byte)0xe7,
                             (byte)0xab, (byte)0xac, (byte)0xd0, (byte)0xd6 };

        public AgileDecryptor()
        {
        }


        public byte[] SecretKey
        {
            get { return _secretKey; }
            set { _secretKey = value; }
        }
        public EncryptionInfo Info
        {
            get { return _info; }
            set { _info = value; }
        }

        public override bool VerifyPassword(string password)
        {
            EncryptionVerifier verifier = _info.Verifier;

            int algorithm = verifier.Algorithm;
            int mode = verifier.CipherMode;

            byte[] pwHash = HashPassword(_info, password);
            byte[] iv = GenerateIv(algorithm, verifier.Salt, null);

            byte[] skey;
            skey = GenerateKey(pwHash, kVerifierInputBlock);
            SymmetricAlgorithm cipher = GetCipher(algorithm, mode, skey, iv);

            byte[] verifierHashInput;

            //using (MemoryStream fStream = new MemoryStream(verifier.Verifier))
            //{
            //    using (CryptoStream cStream = new CryptoStream(fStream, cipher.CreateDecryptor(cipher.Key, cipher.IV),
            //        CryptoStreamMode.Read))
            //    {
            //        verifierHashInput = new byte[cStream.Length];
            //        cStream.Read(verifierHashInput, 0, verifierHashInput.Length);
            //    }
            //}

            verifierHashInput = this.Decrypt(cipher, verifier.Verifier);
            HashAlgorithm sha1 = HashAlgorithm.Create("SHA1");
            byte[] trimmed = new byte[verifier.Salt.Length];
            System.Array.Copy(verifierHashInput, 0, trimmed, 0, trimmed.Length);

            byte[] hashedVerifier = sha1.ComputeHash(trimmed);

            skey = GenerateKey(pwHash, kHashedVerifierBlock);
            iv = GenerateIv(algorithm, verifier.Salt, null);
            cipher = GetCipher(algorithm, mode, skey, iv);
            byte[] verifierHash = Decrypt(cipher, verifier.VerifierHash);
            trimmed = new byte[hashedVerifier.Length];
            System.Array.Copy(verifierHash, 0, trimmed, 0, trimmed.Length);

            if (Arrays.Equals(trimmed, hashedVerifier))
            {
                skey = GenerateKey(pwHash, kCryptoKeyBlock);
                iv = GenerateIv(algorithm, verifier.Salt, null);
                cipher = GetCipher(algorithm, mode, skey, iv);
                byte[] inter = Decrypt(cipher, verifier.EncryptedKey);
                byte[] keySpec = new byte[_info.Header.KeySize / 8];
                Array.Copy(inter, 0, keySpec, 0, keySpec.Length);
                _secretKey = keySpec;
                return true;
            }
            else
                return false;
        }

        public AgileDecryptor(EncryptionInfo info)
        {
            _info = info;
        }

        public override Stream GetDataStream(DirectoryNode dir)
        {
            DocumentInputStream dis = dir.CreateDocumentInputStream("EncryptedPackage");
            _length = dis.ReadLong();
            return new ChunkedCipherInputStream(dis, _length, this);
        }
        public override long Length
        {
            get
            {
                if (_length == -1) throw new InvalidOperationException("EcmaDecryptor.getDataStream() was not called");
                return _length;
            }
        }

        public SymmetricAlgorithm GetCipher(int algorithm, int mode, /*SecretKey key,*/ byte[] key, byte[] vec)
        {
            try
            {
                string name = null;
                string chain = null;
                if (algorithm == EncryptionHeader.ALGORITHM_AES_128 ||
                    algorithm == EncryptionHeader.ALGORITHM_AES_192 ||
                    algorithm == EncryptionHeader.ALGORITHM_AES_256)
                    name = "AES";

                if (mode == EncryptionHeader.MODE_CBC)
                    chain = "CBC";
                else if (mode == EncryptionHeader.MODE_CFB)
                    chain = "CFB";

                //SymmetricAlgorithm cipher = SymmetricAlgorithm.Create(name + "/" + chain + "/Nopadding");
                SymmetricAlgorithm cipher = SymmetricAlgorithm.Create();
                cipher.Key = key;
                cipher.IV = vec;
                cipher.Padding = PaddingMode.None;
                cipher.Mode = chain == "CBC" ? CipherMode.CBC : CipherMode.CFB;
               // cipher.Key = ;
                //cipher.IV = 
                return cipher;
            }
            catch (CryptographicException ex)
            {
                throw ex;
            }
        }
        private byte[] GetBlock(int algorithm, byte[] hash)
        {
            byte[] result = new byte[GetBlockSize(algorithm)];
            for(int i = 0; i < result.Length; i++)
                result[i] = (byte)0x36;
            System.Array.Copy(hash, 0, result, 0, Math.Min(result.Length, hash.Length));
            return result;
        }



        private byte[] GenerateKey(byte[] hash, byte[] blockKey)
        {
            SHA1 sha = new SHA1CryptoServiceProvider();
            //sha1.update(hash); Leon
            //return getBlock(_info.getVerifier().getAlgorithm(), sha1.digest(blockKey));
            byte[] temp = new byte[hash.Length + blockKey.Length];
            Array.Copy(hash, temp, hash.Length);
            Array.Copy(blockKey, 0, temp, hash.Length, blockKey.Length);

            return GetBlock(_info.Verifier.Algorithm, sha.ComputeHash(temp));
        }

        public byte[] GenerateIv(int algorithm, byte[] salt, byte[] blockKey)
        {
            try
            {
                if(blockKey == null)
                    return GetBlock(algorithm, salt);
               // SHA1 sha = new SHA1CryptoServiceProvider();
                //sha.ComputeHash(salt);
                HashAlgorithm sha1 = HashAlgorithm.Create("SHA1");
               // return GetBlock(algorithm, sha
               // sha1.update(salt);
                // return getBlock(algorithm, sha1.digest(blockKey)); Leon
                  sha1.ComputeHash(salt);
             //   sha1 = new SHA1CryptoServiceProvider();
             //   sha1
                return GetBlock(algorithm, sha1.ComputeHash(blockKey));
            }
            catch(System.Security.Cryptography.CryptographicException ex)
            {
                throw ex;
            }
        }

    }

    public class ChunkedCipherInputStream : Stream
    {
        private int _lastIndex = 0;
        private long _pos = 0;
        private long _size;
        private DocumentInputStream _stream;
        private byte[] _chunk;
        private SymmetricAlgorithm _cipher;
        private AgileDecryptor _ag;


        public ChunkedCipherInputStream(DocumentInputStream dis, long size, AgileDecryptor ag)
        {
            try
            {
                _size = size;
                _stream = dis;
                _ag = ag;
                _cipher = _ag.GetCipher(_ag.Info.Header.Algorithm,
                                        _ag.Info.Header.CipherMode,
                                        _ag.SecretKey, _ag.Info.Header.KeySalt);
            }
            catch (System.Security.Cryptography.CryptographicException ex)
            {
                throw ex;
            }
        }

        public int Read()
        {
            byte[] b = new byte[1];
            if (Read(b) == 1)
                return b[0];
            return -1;
        }

        public int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
        }

        public override int Read(byte[] b, int offset, int len)
        {
            int total = 0;

            while (len > 0)
            {
                if (_chunk == null)
                {
                    try
                    {
                        _chunk = NextChunk();
                    }
                    catch (CryptographicException ex)
                    {
                        throw new EncryptedDocumentException(ex.Message);
                    }
                }
                int count = (int)(4096L - (_pos & 0xfff));
                count = Math.Min(Available(), Math.Min(count, len));
                System.Array.Copy(_chunk, (int)(_pos * 0xfff), b, offset, count);
                offset += count;
                len -= count;
                _pos += count;

                if ((_pos & 0xfff) == 0)
                    _chunk = null;
                total += count;
            }

            return total;
        }

        public long Skip(long n)
        {
            long start = _pos;
            long skip = Math.Min(Available(), n);

            if ((((_pos + skip) ^ start) & ~0xfff) != 0)
            {
                _chunk = null;
            }
            _pos += skip;

            return skip;
        }

        public int Available()
        {
            return (int)(_size - _pos);
        }
        public override void Close()
        {
            _stream.Close();
        }

        public bool MarkSupported()
        {
            return false;
        }

        private byte[] NextChunk()
        {
            int index = (int)(_pos >> 12);
            byte[] blockKey = new byte[4];
            LittleEndian.PutInt(blockKey, 0, index);

            byte[] iv = _ag.GenerateIv(_ag.Info.Header.Algorithm,
                                        _ag.Info.Header.KeySalt, blockKey);
            //_cipher.Mode = CipherMode.
            _cipher.Key = _ag.SecretKey;
            _cipher.IV = iv;

            if (_lastIndex != index)
                _stream.Skip((index - _lastIndex) << 12);

            byte[] block = new byte[Math.Min(_stream.Available(), 4096)];
            _stream.ReadFully(block);
            _lastIndex = index + 1;
            throw new NotImplementedException();
            //return _cipher.doFinal(block); Leon
        }

        public override bool CanRead
        {
            get { return _stream.CanRead; }
        }

        public override bool CanSeek
        {
            get { throw new NotImplementedException(); }
        }

        public override bool CanWrite
        {
            get { throw new NotImplementedException(); }
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override long Length
        {
            get { throw new NotImplementedException(); }
        }

        public override long Position
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return _stream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }
    }
}

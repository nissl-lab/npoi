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

using System;
using System.Text;
using System.IO;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI;
using System.Security.Cryptography;

namespace NPOI.POIFS.Crypt
{
    public abstract class Decryptor
    {
        public const string DEFAULT_PASSWORD = "VelvetSweatshop";

        /// <summary>
        /// Return a stream with decrypted data.
        /// Use {@link #getLength()} to get the size of that data that can be safely read from the stream.
        /// Just reading to the end of the input stream is not sufficient because there are
        /// normally padding bytes that must be discarded
        /// </summary>
        /// <param name="dir">the node to read from</param>
        /// <returns>decrypted stream</returns>
        public abstract Stream GetDataStream(DirectoryNode dir);

        public abstract bool VerifyPassword(string password);
     
        /// <summary>
        /// Returns the length of the encytpted data that can be safely read with
        /// {@link #getDataStream(org.apache.poi.poifs.filesystem.DirectoryNode)}.
        /// Just reading to the end of the input stream is not sufficient because there are
        /// normally padding bytes that must be discarded
        /// 
        /// The length variable is initialized in {@link #getDataStream(org.apache.poi.poifs.filesystem.DirectoryNode)},
        /// an attempt to call getLength() prior to getDataStream() will result in IllegalStateException.
        /// 
        /// return length of the encrypted data
        /// </summary>
        public abstract long Length { get; }

        public static Decryptor GetInstance(EncryptionInfo info)
        {
            int major = info.VersionMajor;
            int minor = info.VersionMinor;

            if (major == 4 && minor == 4)
                return new AgileDecryptor(info);
            else if (minor == 2 && (major == 3 || major == 4))
                return new EcmaDecryptor(info);
            else
                throw new EncryptedDocumentException("Unsupported version");
        }

        public Stream GetDataStream(NPOIFSFileSystem fs)
        {
            return GetDataStream(fs.Root);
        }

        public Stream GetDataStream(POIFSFileSystem fs)
        {
            return GetDataStream(fs.Root);
        }

        protected static int GetBlockSize(int algorithm)
        {
            switch (algorithm)
            {
                case EncryptionHeader.ALGORITHM_AES_128:
                    return 16;
                case EncryptionHeader.ALGORITHM_AES_192:
                    return 24;
                case EncryptionHeader.ALGORITHM_AES_256:
                    return 32;
            }
            throw new EncryptedDocumentException("Unknown block size");
        }

        protected byte[] HashPassword(EncryptionInfo info,  string password)
        {
            HashAlgorithm sha1 = HashAlgorithm.Create("SHA1");

            byte[] bytes;
            try
            {
                bytes = Encoding.Unicode.GetBytes(password); //bytes = password.getBytes("UTF-16LE");
            }
            catch (CryptographicUnexpectedOperationException)
            {
                throw new EncryptedDocumentException("UTF16 not supported");
            }

            //sha1.ComputeHash(info.GetVerifier().Salt);
            byte[] salt = info.Verifier.Salt;
            byte[] temp = new byte[salt.Length + bytes.Length];
            Array.Copy(salt, temp, salt.Length);
            Array.Copy(bytes, 0, temp, salt.Length, bytes.Length);
            byte[] hash = sha1.ComputeHash(temp);
            byte[] iterator = new byte[4];
            temp = new byte[24];
            for (int i = 0; i < info.Verifier.SpinCount; i++)
            {
                //sha1.Clear();
                LittleEndian.PutInt(iterator, 0, i);
                //sha1.iterator; //sha1.update(iterator);
                Array.Copy(iterator, temp, iterator.Length);
                Array.Copy(hash, 0, temp, iterator.Length, hash.Length);
                hash = sha1.ComputeHash(temp);
            }

            return hash;
        }

        protected byte[] Decrypt(SymmetricAlgorithm cipher,byte[] encryptBytes)
        {
            byte[] decryptBytes = new byte[0];
            using (MemoryStream fStream = new MemoryStream(encryptBytes))
            {
                using (CryptoStream cStream = new CryptoStream(fStream, cipher.CreateDecryptor(cipher.Key, cipher.IV),
                    CryptoStreamMode.Read))
                {
                    
                    using (MemoryStream destMs = new MemoryStream())
                    {
                        byte[] buffer = new byte[100];
                        int readLen;

                        while ((readLen = cStream.Read(buffer, 0, 100)) > 0)
                            destMs.Write(buffer, 0, readLen);

                        decryptBytes = destMs.ToArray();
                    }
                }
            }
            return decryptBytes;
        }
    }
}

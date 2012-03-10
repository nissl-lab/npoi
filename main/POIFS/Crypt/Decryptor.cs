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
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI;
using System.Security;
using System.Security.Cryptography;

namespace NPOI.POIFS.Crypt
{
    public abstract class Decryptor
    {
        public static string DEFAULT_PASSWORD = "VelvetSweatshop";

        public abstract Stream GetDataStream(DirectoryNode dir);

        public abstract bool VerifyPassword(string password);

        public static Decryptor GetInstance(EncryptionInfo info)
        {
            int major = info.GetVersionMajor();
            int minor = info.GetVersionMinor();

            if (major == 4 && minor == 4)
                return new AgileDecryptor(info);
            else if (minor == 2 && (major == 3 || major == 4))
                return new EcmaDecryptor(info);
            else
                throw new EncryptedDocumentException("Unsupported version");
        }

        public Stream GetDataStream(NPOIFSFileSystem fs)
        {
            return GetDataStream(fs.GetRoot());
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
                bytes = Encoding.UTF7.GetBytes(password); //bytes = password.getBytes("UTF-16LE");
            }
            catch (CryptographicUnexpectedOperationException ex)
            {
                throw new EncryptedDocumentException("UTF16 not supported");
            }

            //sha1.ComputeHash(info.GetVerifier().Salt);
            byte[] hash = sha1.ComputeHash(bytes);
            byte[] iterator = new byte[4];

            for (int i = 0; i < info.GetVerifier().SpinCount; i++)
            {
                sha1.Clear();
                LittleEndian.PutInt(iterator, i);
                //sha1.iterator; //sha1.update(iterator);
                hash = sha1.ComputeHash(hash);
            }

            return hash;
        }
    }
}

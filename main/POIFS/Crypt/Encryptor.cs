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

using System.IO;
using NPOI.POIFS.FileSystem;

namespace NPOI.POIFS.Crypt
{
    using NPOI.Util;
    using System;
    using System.Collections.Generic;

    public interface IKey
    {
        string GetAlgorithm();

        string GetFormat();

        byte[] GetEncoded();
    }

    public interface ISecretKey: IKey
    {

    }
    public interface IPrivateKey: IKey
    {

    }
    public abstract class Encryptor
    {
        public static readonly string DEFAULT_POIFS_ENTRY = Decryptor.DEFAULT_POIFS_ENTRY;

        protected EncryptionInfo _encryptionInfo;
        private ISecretKey _secretKey;

        protected Encryptor() { }

        protected Encryptor(Encryptor other)
        {
            _encryptionInfo = other._encryptionInfo;
            // SecretKey is immutable by design in POI; share reference.
            _secretKey = other._secretKey;
        }

        /// <summary>
        /// Return an output stream for encrypted data (writes to the given POIFS directory node).
        /// </summary>
        public abstract OutputStream GetDataStream(DirectoryNode dir);

        /// <summary>
        /// For tests / algorithm setup: write-side password confirmation and material setup.
        /// </summary>
        public abstract void ConfirmPassword(
            string password,
            byte[] keySpec,
            byte[] keySalt,
            byte[] verifier,
            byte[] verifierSalt,
            byte[] integritySalt);

        public abstract void ConfirmPassword(string password);

        /// <summary>
        /// Factory: obtain encryptor from EncryptionInfo.
        /// </summary>
        public static Encryptor GetInstance(EncryptionInfo info) => info.Encryptor;

 		public OutputStream GetDataStream(NPOIFSFileSystem fs)
		{
			return GetDataStream(fs.Root);
		}
		
        public OutputStream GetDataStream(OPOIFSFileSystem fs)
        {
            return GetDataStream(fs.Root);
        }
        /// <summary>
        /// Optional: return a stream that writes encrypted bytes directly to the given stream.
        /// Default: not supported (matches Java behavior).
        /// </summary>
        public virtual ChunkedCipherOutputStream GetDataStream(OutputStream stream, int initialOffset)
        {
            throw new EncryptedDocumentException("this decryptor doesn't support writing directly to a stream");
        }

        public ISecretKey GetSecretKey() => _secretKey;

        public void SetSecretKey(ISecretKey secretKey) => _secretKey = secretKey;

        public EncryptionInfo GetEncryptionInfo() => _encryptionInfo;

        public void SetEncryptionInfo(EncryptionInfo encryptionInfo) => _encryptionInfo = encryptionInfo;

        /// <summary>
        /// Sets the chunk size of the data stream.
        /// Needs to be set before the data stream is requested.
        /// Default: not supported (implementation-specific).
        /// </summary>
        public virtual void SetChunkSize(int chunkSize)
        {
            throw new EncryptedDocumentException("this decryptor doesn't support changing the chunk size");
        }

        /// <summary>
        /// Generic record-like properties for diagnostics/inspection.
        /// Mirrors the Java GenericRecord return shape with secret key bytes.
        /// </summary>
        public virtual IDictionary<string, Func<object>> GetGenericProperties()
        {
            return new Dictionary<string, Func<object>>
            {
                { "secretKey", () => _secretKey == null ? null : (object)_secretKey.GetEncoded() }
            };
        }
    }
}
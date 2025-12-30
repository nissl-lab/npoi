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

namespace NPOI.POIFS.Crypt
{
    using System;

    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    public abstract class Decryptor
    {
        public static String DEFAULT_PASSWORD = "VelvetSweatshop";
        public static String DEFAULT_POIFS_ENTRY = "EncryptedPackage";

        protected internal IEncryptionInfoBuilder builder;
        private ISecretKey secretKey;
        private byte[] verifier, integrityHmacKey, integrityHmacValue;

        protected Decryptor(IEncryptionInfoBuilder builder)
        {
            this.builder = builder;
        }

        /**
     * Return a stream with decrypted data.
     * <p>
     * Use {@link #getLength()} to Get the size of that data that can be safely read from the stream.
     * Just Reading to the end of the input stream is not sufficient because there are
     * normally pAdding bytes that must be discarded
     * </p>
     *
     * @param dir the node to read from
     * @return decrypted stream
     */
        public abstract InputStream GetDataStream(DirectoryNode dir);

        /**
         * Wraps a stream for decryption<p>
         *
         * As we are handling streams and don't know the total length beforehand,
         * it's the callers duty to care for the length of the entries.
         *
         * @param stream the stream to be wrapped
         * @param initialPos initial/current byte position within the stream
         * @return decrypted stream
         */
        public virtual InputStream GetDataStream(InputStream stream, int size, int initialPos)
        {
            throw new EncryptedDocumentException("this decryptor doesn't support reading from a stream");
        }

        public abstract bool VerifyPassword(String password);

        /**
     * Returns the length of the encrypted data that can be safely read with
     * {@link #getDataStream(NPOI.POIFS.FileSystem.DirectoryNode)}.
     * Just Reading to the end of the input stream is not sufficient because there are
     * normally pAdding bytes that must be discarded
     *
     * <p>
     *    The length variable is Initialized in {@link #getDataStream(NPOI.POIFS.FileSystem.DirectoryNode)},
     *    an attempt to call GetLength() prior to GetDataStream() will result in InvalidOperationException.
     * </p>
     *
     * @return length of the encrypted data
     * @throws InvalidOperationException if {@link #getDataStream(NPOI.POIFS.FileSystem.DirectoryNode)}
     * was not called
     */
        public abstract long GetLength();


        /**
         * Sets the chunk size of the data stream.
         * Needs to be set before the data stream is requested.
         * When not set, the implementation uses method specific default values
         *
         * @param chunkSize the chunk size, i.e. the block size with the same encryption key
         */
        public virtual void SetChunkSize(int chunkSize)
        {
            throw new EncryptedDocumentException("this decryptor doesn't support changing the chunk size");
        }

        public static Decryptor GetInstance(EncryptionInfo info)
        {
            Decryptor d = info.Decryptor;
            if (d == null)
            {
                throw new EncryptedDocumentException("Unsupported version");
            }
            return d;
        }

        public InputStream GetDataStream(NPOIFSFileSystem fs)
        {
            return GetDataStream(fs.Root);
        }
        public InputStream GetDataStream(OPOIFSFileSystem fs)
        {
            return GetDataStream(fs.Root);
        }
        public InputStream GetDataStream(POIFSFileSystem fs)
        {
            return GetDataStream(fs.Root);
        }

        // for tests
        public byte[] GetVerifier()
        {
            return verifier;
        }

        public ISecretKey GetSecretKey()
        {
            return secretKey;
        }

        public byte[] GetIntegrityHmacKey()
        {
            return integrityHmacKey;
        }

        public byte[] GetIntegrityHmacValue()
        {
            return integrityHmacValue;
        }

        protected void SetSecretKey(ISecretKey secretKey)
        {
            this.secretKey = secretKey;
        }

        protected void SetVerifier(byte[] verifier)
        {
            this.verifier = verifier;
        }

        protected void SetIntegrityHmacKey(byte[] integrityHmacKey)
        {
            this.integrityHmacKey = integrityHmacKey;
        }

        protected void SetIntegrityHmacValue(byte[] integrityHmacValue)
        {
            this.integrityHmacValue = integrityHmacValue;
        }

        protected int GetBlockSizeInBytes()
        {
            return builder.GetHeader().BlockSize;
        }

        protected int GetKeySizeInBytes()
        {
            return builder.GetHeader().KeySize/8;
        }
    }
}
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
namespace NPOI.POIFS.Crypt
{
    using System;
    using System.IO;
    using NPOI.POIFS.EventFileSystem;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using Org.BouncyCastle.Crypto;
    using SixLabors.ImageSharp.Memory;
    using System.Security.Cryptography;

    public abstract class ChunkedCipherOutputStream : LittleEndianOutputStream
    {
        private const int STREAMING = -1;

        private readonly int _chunkSize;
        private readonly int _chunkBits;

        private readonly byte[] _chunk;
        private readonly SparseBitSet _plainByteFlags;
        private readonly FileInfo fileOut;
        private readonly DirectoryNode dir;

        private long _pos;
        private long _totalPos;
        private long _written;

        private Cipher _cipher;
        private bool _isClosed;
        protected readonly IEncryptionInfoBuilder builder;
        protected readonly Encryptor encryptor;

        protected ChunkedCipherOutputStream(OutputStream stream, int chunkSize, IEncryptionInfoBuilder builder, Encryptor encryptor) : base(stream)
        {
            _chunkSize = chunkSize;
            int cs = (chunkSize == STREAMING) ? 4096 : chunkSize;
            _chunk = new byte[cs];
            _plainByteFlags = new SparseBitSet(cs);
            _chunkBits =  Number.BitCount(GetChunkMask());
            this.fileOut = null;
            this.dir = null;
            this.builder = builder;
            this.encryptor = encryptor;
            _cipher = InitCipherForBlock(null, 0, lastChunk: false);
        }

        protected ChunkedCipherOutputStream(DirectoryNode dir, int chunkSize, IEncryptionInfoBuilder builder, Encryptor encryptor) : base(null)
        {
            _chunkSize = chunkSize;
            int cs = chunkSize == STREAMING ? 4096 : chunkSize;
            _chunk = new byte[chunkSize];
            _plainByteFlags = new SparseBitSet(cs);
            _chunkBits = Number.BitCount(GetChunkMask());
            fileOut = TempFile.CreateTempFile("encrypted_package", "crypt");
            this.out1 = new FileStream(fileOut.FullName, FileMode.Create, FileAccess.ReadWrite, FileShare.Read);
            this.dir = dir;
            this.builder = builder;
            this.encryptor = encryptor;
            _cipher = InitCipherForBlock(null, 0, false);
        }

        public Cipher InitCipherForBlock(int block, bool lastChunk)
            => InitCipherForBlock(_cipher, block, lastChunk);

        protected abstract Cipher InitCipherForBlock(Cipher existing, int block, bool lastChunk);

        protected abstract void CalculateChecksum(FileInfo fileOut, int oleStreamSize);

        protected abstract void CreateEncryptionInfoEntry(DirectoryNode dir, FileInfo tmpFile);

        protected virtual Cipher InitCipherForBlockNoFlush(Cipher existing, int block, bool lastChunk)
            => InitCipherForBlock(existing, block, lastChunk);

        public long Length => _written;
        public long Position
        {
            get => _totalPos;
            set => throw new NotSupportedException();
        }

        public override void Write(int b)
        {
            Write(new byte[] { (byte)b });
        }

        public override void Write(byte[] b)
        {
            Write(b, 0, b.Length);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            WriteInternal(buffer, offset, count, writePlain: false);
        }

        public void WritePlain(byte[] buffer, int offset, int count)
        {
            WriteInternal(buffer, offset, count, writePlain: true);
        }

        private void WriteInternal(byte[] buffer, int offset, int count, bool writePlain)
        {
            if (count == 0)
            {
                return;
            }

            if (count < 0 || buffer.Length < offset + count)
            {
                throw new IOException("not enough bytes in your input buffer");
            }

            int chunkMask = GetChunkMask();
            while (count > 0)
            {
                int posInChunk = (int)(_pos & chunkMask);
                int spaceInChunk = _chunk.Length - posInChunk;
                int nextLen = Math.Min(spaceInChunk, count);

                Buffer.BlockCopy(buffer, offset, _chunk, posInChunk, nextLen);
                if (writePlain)
                {
                    _plainByteFlags.Set(posInChunk, posInChunk + nextLen);
                }

                _pos += nextLen;
                _totalPos += nextLen;
                offset += nextLen;
                count -= nextLen;

                if ((_pos & chunkMask) == 0)
                {
                    // chunk boundary reached
                    WriteChunk(continued: count > 0);
                }
            }
        }

        protected int GetChunkMask() => _chunk.Length - 1;

        protected void WriteChunk(bool continued)
        {
            if (_pos == 0 || _totalPos == _written)
                return;

            int posInChunk = (int)(_pos & GetChunkMask());
            int index = (int)(_pos >> _chunkBits);

            bool lastChunk;
            if (posInChunk == 0)
            {
                index--;
                posInChunk = _chunk.Length;
                lastChunk = false;
            }
            else
            {
                lastChunk = true; // pad last chunk
            }

            int ciLen;
            try
            {
                bool doFinal = true;
                long oldPos = _pos;

                _pos = 0; // reset (streaming scenario)
                if (_chunkSize == STREAMING)
                {
                    if (continued)
                        doFinal = false;
                }
                else
                {
                    _cipher = InitCipherForBlock(_cipher, index, lastChunk);
                    _pos = oldPos; // restore for non-streaming
                }

                ciLen = InvokeCipher(posInChunk, doFinal);
            }
            catch (CryptographicException e)
            {
                throw new IOException("Can't (re)initialize cipher.", e);
            }

            out1.Write(_chunk, 0, ciLen);
            _plainByteFlags.Clear();
            _written += ciLen;
        }

        /// <summary>
        /// Allows XOR-like schemes to override cipher invocation.
        /// Default uses <see cref="Cipher"/>.
        /// </summary>
        protected virtual int InvokeCipher(int posInChunk, bool doFinal)
        {
            byte[] plain = _plainByteFlags.IsEmpty ? null : (byte[])_chunk.Clone();

            int ciLen;
            if (doFinal)
            {
                // Final block for this (sub)chunk: run DoFinal (may add/remove padding depending on mode)
                // BouncyCastle allows in-place finalization when input == output.
                ciLen = _cipher.DoFinal(_chunk, 0, posInChunk, _chunk);

                /*
                 * Re-initialize cipher for the next logical block.
                 * Original logic relied on (posInChunk == 0) meaning "full chunk" which is misleading.
                 * At this point a full chunk arrives here with posInChunk == _chunk.Length.
                 */
                if (_cipher.Algorithm == CipherAlgorithm.rc4
                    && _cipher.Algorithm.provider.cipherProviderName.ToUpperInvariant() == "IBMJCE")
                {
                    int index = (int)(_pos >> _chunkBits);
                    bool lastChunk = (posInChunk != 0);
                    if (posInChunk == 0)
                    {
                        index--;
                        posInChunk = _chunk.Length;
                        lastChunk = false;
                    }

                    _cipher = InitCipherForBlockNoFlush(_cipher, index, lastChunk);
                }
            }
            else
            {
                // Intermediate portion: Update only (no padding, length == returned bytes)
                ciLen = _cipher.Update(_chunk, 0, posInChunk, _chunk, 0);
            }

            if (plain != null)
            {
                int i = _plainByteFlags.NextSetBit(0);
                while (i >= 0 && i < posInChunk)
                {
                    _chunk[i] = plain[i];
                    i = _plainByteFlags.NextSetBit(i + 1);
                }
            }

            return ciLen;
        }

        /// <summary>
        /// Some ciphers (e.g., XOR) are based on the record size, which needs to be set before encryption.
        /// </summary>
        public virtual void SetNextRecordSize(int recordSize, bool isPlain)
        {
            // default no-op
        }

        public override void Close()
        {
            if (_isClosed)
                return;
            _isClosed = true;

            try
            {
                WriteChunk(continued: false);

                if (fileOut != null)
                {
                    //Ensure close the file so the file lock is released
                    out1.Close();
                    
                    int oleStreamSize = (int)(fileOut.Length + LittleEndianConsts.LONG_SIZE);
                    CalculateChecksum(fileOut, (int)_pos);
                    dir.CreateDocument(Decryptor.DEFAULT_POIFS_ENTRY, oleStreamSize, new EncryptedPackageWriter(this));
                    CreateEncryptionInfoEntry(dir, fileOut);
                }
            }
            catch (Exception e)
            {
                throw new IOException(e.Message);
            }
            finally
            {
                base.Close();
                // safe after disposting stream in base class
                if (fileOut != null)
                    fileOut.Delete();
            }
        }

        // ---- Small helper "sparse" bitset (range marks for plaintext bytes) ----
        protected sealed class SparseBitSet
        {
            private readonly System.Collections.BitArray _bits;

            public SparseBitSet(int length) => _bits = new System.Collections.BitArray(length, false);

            public void Set(int fromInclusive, int toExclusive)
            {
                if (fromInclusive < 0)
                    fromInclusive = 0;
                if (toExclusive > _bits.Length)
                    toExclusive = _bits.Length;
                for (int i = fromInclusive; i < toExclusive; i++)
                    _bits[i] = true;
            }

            public void Clear() => _bits.SetAll(false);

            public bool IsEmpty
            {
                get
                {
                    for (int i = 0; i < _bits.Length; i++)
                        if (_bits[i])
                            return false;
                    return true;
                }
            }

            public int NextSetBit(int fromIndex)
            {
                for (int i = Math.Max(0, fromIndex); i < _bits.Length; i++)
                    if (_bits[i])
                        return i;
                return -1;
            }
        }

        private sealed class EncryptedPackageWriter : POIFSWriterListener {
            readonly ChunkedCipherOutputStream stream;
            public EncryptedPackageWriter(ChunkedCipherOutputStream stream)
            {
                this.stream = stream;
            }
            public void ProcessPOIFSWriterEvent(POIFSWriterEvent event1) {
                try {
                    DocumentOutputStream os = event1.Stream;
                    byte[] buf = new byte[stream._chunkSize];

                    // StreamSize (8 bytes): An unsigned integer that specifies the number of bytes used by data 
                    // encrypted within the EncryptedData field, not including the size of the StreamSize field. 
                    // Note that the actual size of the \EncryptedPackage stream (1) can be larger than this 
                    // value, depending on the block size of the chosen encryption algorithm
                    LittleEndian.PutLong(buf, 0, stream._pos);
                    os.Write(buf, 0, LittleEndian.LONG_SIZE);

                    // Open existing file for READ (do not truncate)
                    using (var fis = new FileStream(stream.fileOut.FullName, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        int read;
                        while ((read = fis.Read(buf, 0, buf.Length)) > 0)
                        {
                            os.Write(buf, 0, read);
                        }
                    }
                    
                    os.Close();

                } catch (IOException e) {
                    throw new EncryptedDocumentException(e);
                }
            }
        }
    }

}
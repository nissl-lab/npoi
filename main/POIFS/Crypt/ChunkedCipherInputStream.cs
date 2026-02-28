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
using System.Security;
using NPOI.Util;
using System;
using System.Buffers;
using System.Security.Cryptography;

namespace NPOI.POIFS.Crypt
{
    public abstract class ChunkedCipherInputStream : LittleEndianInputStream
    {
        private const int STREAMING = -1;

        private readonly Stream _baseStream;

        private readonly int _chunkSize;
        private readonly int _chunkBits;
        private readonly long _size;

        private readonly byte[] _chunk;
        private readonly byte[] _plain;

        private Cipher _cipher;
        private int _lastIndex;
        private long _pos;
        private bool _chunkIsValid;
        protected IEncryptionInfoBuilder builder;
        protected Decryptor decryptor;

        protected ChunkedCipherInputStream(InputStream stream, long size, int chunkSize, Decryptor decryptor)
            : this(stream, size, chunkSize, 0, decryptor)
        { }

        protected ChunkedCipherInputStream(InputStream stream, long size, int chunkSize, int initialPos, Decryptor decryptor)
            : base(stream)
        {
            _baseStream = stream ?? throw new ArgumentNullException(nameof(stream));
            
            if(!stream.CanRead)
                throw new ArgumentException("Stream must be readable.", nameof(stream));

            this._size = size;
            this._pos = initialPos;
            this._chunkSize = chunkSize;

            this.builder = decryptor.builder;
            this.decryptor = decryptor;

            var cs = chunkSize == -1 ? 4096 : chunkSize;

            this._chunk = IOUtils.SafelyAllocate(cs, CryptoFunctions.MAX_RECORD_LENGTH);
            this._plain = IOUtils.SafelyAllocate(cs, CryptoFunctions.MAX_RECORD_LENGTH);
            this._chunkBits = Number.BitCount(_chunk.Length - 1);
            this._lastIndex = (int)(_pos >> _chunkBits);
            this._cipher = InitCipherForBlock(null, _lastIndex);
        }

        public Cipher InitCipherForBlock(int block)
        {
            if(_chunkSize != STREAMING)
                throw new CryptographicException("The cipher block can only be set for streaming encryption (chunkSize == -1).");

            _chunkIsValid = false;
            _cipher = InitCipherForBlock(_cipher, block);
            return _cipher;
        }

        /// <summary>
        /// Implement this in subclasses to (re)initialize the transform for a given block index.
        /// Return <paramref name="existing"/> if you can re-use it, otherwise return a new transform.
        /// </summary>
        protected abstract Cipher InitCipherForBlock(Cipher existing, int block);

        #region Stream overrides

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;

        public override long Length => _size;

        public override long Position
        {
            get => _pos;
            set => _pos = value;
        }

        public override void Flush() { /* no-op */ }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override int ReadByte()
        {
            return (byte)ReadUByte();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return ReadInternal(buffer, offset, count, readPlain: false);
        }

        #endregion

        /// <summary>
        /// Used when BIFF header fields are being read. The internal transform must step
        /// even when unencrypted bytes are read.
        /// </summary>
        public void ReadPlain(byte[] buffer, int offset, int count)
        {
            if(count <= 0)
                return;

            int total = 0;
            while(total < count)
            {
                int read = ReadInternal(buffer, offset + total, count - total, readPlain: true);
                if(read <= 0)
                    throw new EndOfStreamException("buffer underrun");
                total += read;
            }
        }

        /// <summary>
        /// Some ciphers (e.g., XOR) are based on the record size, which needs to be set before decryption.
        /// Default no-op.
        /// </summary>
        public virtual void SetNextRecordSize(int recordSize) { }

        /// <summary>Returns the internal chunk buffer (encrypted)</summary>
        protected byte[] GetChunk() => _chunk;

        /// <summary>Returns the internal plain buffer (pre-encryption copy)</summary>
        protected byte[] GetPlain() => _plain;

        /// <summary>Absolute position in the logical stream</summary>
        public long GetPos() => _pos;

        // ---- Internal helpers ----

        private int ReadInternal(byte[] destination, int offset, int count, bool readPlain)
        {
            if(RemainingBytes() <= 0)
                return -1;

            int total = 0;
            int chunkMask = GetChunkMask();

            while(count > 0)
            {
                if(!_chunkIsValid)
                {
                    try
                    {
                        NextChunk();
                        _chunkIsValid = true;
                    }
                    catch(CryptographicException e)
                    {
                        throw new IOException("Cipher error", e);
                    }
                }

                int posInChunk = (int)(_pos & chunkMask);
                int availInChunk = _chunk.Length - posInChunk;
                int availStream = RemainingBytes();
                if(availStream == 0)
                    return total;

                int toCopy = Math.Min(availStream, Math.Min(availInChunk, count));
                byte[] src = readPlain ? _plain : _chunk;
                Buffer.BlockCopy(src, posInChunk, destination, offset, toCopy);

                offset += toCopy;
                count -= toCopy;
                _pos += toCopy;
                total += toCopy;

                if((_pos & chunkMask) == 0)
                {
                    _chunkIsValid = false;
                }
            }

            return total;
        }

        private int GetChunkMask() => _chunk.Length - 1;

        public override int Available()
        {
            return RemainingBytes();
        }

        private int RemainingBytes()
        {
            return (int)(_size - _pos);
        }

        private void NextChunk()
        {
            if(_chunkSize != STREAMING)
            {
                int index = (int)(_pos >> _chunkBits);
                _cipher = InitCipherForBlock(_cipher, index);

                if(_lastIndex != index)
                {
                    long skipN = ((long)index - _lastIndex) << _chunkBits;
                    if(SkipUnderlying(_baseStream, skipN) < skipN)
                        throw new EndOfStreamException("buffer underrun");
                }
                _lastIndex = index + 1;
            }

            int todo = (int)Math.Min(_size, _chunk.Length);

            int readBytes, totalBytes = 0;
            do
            {
                readBytes = _baseStream.Read(_plain, totalBytes, todo - totalBytes);
                if(readBytes > 0)
                    totalBytes += readBytes;
            } while(readBytes != 0 && totalBytes < todo);

            if(readBytes == 0 && _pos + totalBytes < _size && _size < int.MaxValue)
                throw new EndOfStreamException("buffer underrun");

            // Encrypted data is processed in 16-byte multiples; pull a little more if needed
            if(totalBytes % 16 != 0)
            {
                int toRead = 16 - (totalBytes % 16);
                int add = _baseStream.Read(_plain, totalBytes, toRead);
                if(add > 0)
                    totalBytes += add;
            }

            Buffer.BlockCopy(_plain, 0, _chunk, 0, totalBytes);
            InvokeCipher(totalBytes, doFinal: totalBytes == _chunkSize);
        }

        /// <summary>
        /// Allows XOR-like schemes to override cipher invocation.
        /// Default uses <see cref="Cipher"/>.
        /// </summary>
        protected virtual int InvokeCipher(int totalBytes, bool doFinal)
        {
            if(_cipher == null || totalBytes == 0)
                return 0;

            if(doFinal)
            {
                // Process and finalize the current chunk
                return _cipher.DoFinal(_chunk, 0, totalBytes, _chunk);
            }
            else
            {
                // Intermediate chunk, no padding applied/removed
                return _cipher.Update(_chunk, 0, totalBytes, _chunk, 0);
            }
        }

        private static long SkipUnderlying(Stream s, long n)
        {
            if(n <= 0)
                return 0;
            if(s.CanSeek)
            {
                long cur = s.Position;
                long len = s.Length;
                long target = Math.Min(len, cur + n);
                s.Position = target;
                return target - cur;
            }

            const int BufSize = 8192;
            byte[] buf = ArrayPool<byte>.Shared.Rent(BufSize);
            long remaining = n;
            try
            {
                while(remaining > 0)
                {
                    int toRead = (int)Math.Min(BufSize, remaining);
                    int r = s.Read(buf, 0, toRead);
                    if(r <= 0)
                        break;
                    remaining -= r;
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buf);
            }
            return n - remaining;
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }
    }
}
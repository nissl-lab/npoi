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

namespace NPOI.POIFS.Crypt
{
    using System;
    using System.Collections;

    public abstract class ChunkedCipherInputStream : LittleEndianInputStream
    {
        private readonly int _chunkSize;
        private readonly int _chunkBits;

        private readonly long _size;
        private readonly byte[] _chunk;
        private readonly byte[] _plain;
        private readonly Cipher _cipher;

        private int _lastIndex = 0;
        private long _pos = 0;
        private bool _chunkIsValid;

        protected IEncryptionInfoBuilder builder;
        protected Decryptor decryptor;

        protected ChunkedCipherInputStream(
            InputStream stream,
            long size,
            int chunkSize,
            IEncryptionInfoBuilder builder,
            Decryptor decryptor)
            : this(stream, size, chunkSize, 0, builder, decryptor)
        { }

        protected ChunkedCipherInputStream(
            InputStream stream,
            long size,
            int chunkSize,
            int initialPos,
            IEncryptionInfoBuilder builder,
            Decryptor decryptor)
            : base(stream)
        {
            this._size = size;
            this._pos = initialPos;
            this._chunkSize = chunkSize;

            this.builder = builder;
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
            if (_chunkSize != -1)
            {
                throw new SecurityException("the cipher block can only be set for streaming encryption, e.g. CryptoAPI...");
            }

            _chunkIsValid = false;
            return InitCipherForBlock(_cipher, block);
        }

        protected abstract Cipher InitCipherForBlock(Cipher existing, int block);

        public override int Read()
        {
            byte[] b = new byte[1];
            if (Read(b) == 1)
                return b[0] & 0xFF;
            return -1;
        }

        // do not implement! -> recursion
        // public int Read(byte[] b) throws IOException;

        public override int Read(byte[] b, int off, int len)
        {
            return Read(b, off, len, false);
        }

        public int Read(byte[] b, int off, int len, bool readPlain)
        {
            int total = 0;

            if (RemainingBytes() <= 0) return 0;

            int chunkMask = GetChunkMask();

            while (len > 0)
            {
                if (!_chunkIsValid)
                {
                    try
                    {
                        NextChunk();
                        _chunkIsValid = true;
                    }
                    catch (SecurityException e)
                    {
                        throw new EncryptedDocumentException(e.Message, e);
                    }
                }
                int count = (int)(_chunk.Length - (_pos & chunkMask));
                int avail = RemainingBytes();
                if (avail == 0)
                {
                    return total;
                }
                count = Math.Min(avail, Math.Min(count, len));

                Array.Copy(readPlain ? _plain : _chunk, (int)(_pos & chunkMask), b, off, count);

                off += count;
                len -= count;
                _pos += count;
                if ((_pos & chunkMask) == 0)
                {
                    _chunkIsValid = false;
                }
                total += count;
            }

            return total;
        }

        public override long Skip(long n)
        {
            long start = _pos;
            long skip = Math.Min(RemainingBytes(), n);

            if ((((_pos + skip) ^ start) & ~GetChunkMask()) != 0)
            {
                _chunkIsValid = false;
            }

            _pos += skip;
            return skip;
        }

        public override int Available()
        {
            return RemainingBytes();
        }

        /// <summary>
        /// Helper method for forbidden available call - we know the size beforehand, so it's ok ...
        /// </summary>
        /// <returns>the remaining byte until EOF</returns>
        private int RemainingBytes()
        {
            return (int)(_size - _pos);
        }

        public override bool MarkSupported()
        {
            return false;
        }

        public override void Mark(int readlimit)
        {
            throw new InvalidOperationException();
        }

        public override void Reset()
        {
            throw new InvalidOperationException();
        }

        protected int GetChunkMask()
        {
            return _chunk.Length - 1;
        }

        private void NextChunk()
        {
            if (_chunkSize != 0)
            {
                int index = (int)(_pos >> _chunkBits);
                InitCipherForBlock(_cipher, index);

                if (_lastIndex != index)
                {
                    long skipN = (index - _lastIndex) << _chunkBits;
                    if (base.Skip(skipN) < skipN)
                    {
                        throw new EndOfStreamException("buffer underrun");
                    }
                }

                _lastIndex = index + 1;
            }

            int todo = (int)Math.Min(_size, _chunk.Length);
            int readBytes, totalBytes = 0;
            do
            {
                readBytes = base.Read(_plain, totalBytes, todo - totalBytes);
                totalBytes += Math.Max(0, readBytes);
            } while (readBytes != 0 && totalBytes < todo);

            if (readBytes == 0 && _pos + totalBytes < _size && _size < int.MaxValue)
            {
                throw new EndOfStreamException("buffer underrun");
            }

            Array.Copy(_plain, 0, _chunk, 0, totalBytes);

            InvokeCipher(totalBytes, totalBytes == _chunkSize);
        }

        /// <summary>
        /// Helper function for overriding the cipher invocation, i.e. XOR doesn't use a cipher and uses its own implementation
        /// </summary>
        /// <param name="totalBytes">The total bytes.</param>
        /// <param name="doFinal">The do final.</param>
        /// <returns></returns>
        protected int InvokeCipher(int totalBytes, bool doFinal)
        {
            if (doFinal)
            {
                return _cipher.DoFinal(_chunk, 0, totalBytes, _chunk);
            }
            else
            {
                return _cipher.Update(_chunk, 0, totalBytes, _chunk);
            }
        }

        /// <summary>
        /// Used when BIFF header fields (sid, size) are being read. The internal <see cref="Cipher"/> instance must step even when unencrypted bytes are read
        /// </summary>
        /// <param name="b">The buffet.</param>
        /// <param name="off">The offset.</param>
        /// <param name="len">The length.</param>
        /// <exception cref="System.IO.EndOfStreamException">buffer underrun</exception>
        /// <exception cref="NPOI.Util.RuntimeException"></exception>
        public void ReadPlain(byte[] b, int off, int len)
        {
            if (len <= 0)
            {
                return;
            }

            try
            {
                int readBytes, total = 0;
                do
                {
                    readBytes = Read(b, off, len, true);
                    total += Math.Max(0, readBytes);
                } while (readBytes > 0 && total < len);

                if (total < len)
                {
                    throw new EndOfStreamException("buffer underrun");
                }
            }
            catch (IOException e)
            {
                // need to wrap checked exception, because of LittleEndianInput interface :(
                throw new RuntimeException(e);
            }
        }

        /// <summary>
        /// Some ciphers (actually just XOR) are based on the record size, which needs to be set before decryption
        /// </summary>
        /// <param name="recordSize">The size of the next record.</param>
        public void SetNextRecordSize(int recordSize) { }

        /// <summary>
        /// Gets the chunk bytes.
        /// </summary>
        /// <returns>the chunk bytes</returns>
        protected byte[] GetChunk()
        {
            return _chunk;
        }

        /// <summary>
        /// Gets the plain bytes.
        /// </summary>
        /// <returns>the plain bytes</returns>
        protected byte[] GetPlain()
        {
            return _plain;
        }

        /// <summary>
        /// Gets the position.
        /// </summary>
        /// <returns>the absolute position in the stream</returns>
        public long GetPos()
        {
            return _pos;
        }
    }

}